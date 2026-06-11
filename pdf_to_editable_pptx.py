#!/usr/bin/env python3
"""
pdf_to_editable_pptx.py
=======================

WHAT THIS SCRIPT DOES (the big picture)
---------------------------------------
You give it a PDF (like a workbook lesson). It builds a PowerPoint file where:

  1. EVERY PIECE OF TEXT becomes a real, editable PowerPoint text box,
     placed at the same position, size, and color as in the PDF.

  2. SIMPLE LAYOUT GRAPHICS (colored boxes, header bands, table grids,
     answer lines) become real PowerPoint shapes you can move and restyle.

  3. DECORATIVE ARTWORK (cartoon characters, icons, stars, speech bubbles)
     becomes individual MOVABLE PICTURES with transparent surroundings.
     Each one is layered above the colored boxes (so a character can lean
     over a box, like the original) but below the text - and because each
     is a normal picture, you can click and drag it out of the way
     whenever it overlaps something you're editing. Text that lives
     INSIDE artwork (like the words in a speech bubble) travels with its
     picture instead of becoming editable text.

  4. The PAGE'S BACKGROUND TEXTURES (colored borders, header bands) are
     baked into each slide's locked background image, so the page keeps
     its look without loose pieces to bump.

  5. WATERMARK FOOTER LINES are removed automatically: any line of text
     containing a phrase on the removal list (by default "Downloaded by"
     and "Copying is not permitted") is left out of the output, along
     with the little strip drawn behind it.

HOW IT TELLS "ARTWORK" APART FROM "LAYOUT"
------------------------------------------
Inside a PDF, everything on the page is one of three things:
  * text        (letters with a font, size, and position)
  * drawings    (vector shapes: lines, rectangles, and curves)
  * images      (embedded pictures, like photos or textures)

Cartoon characters are built from MANY SMALL shapes (lots of them curved)
packed closely together. Layout elements are the opposite: a FEW LARGE
shapes with mostly straight edges (a box, a line, a table border). The
script uses that difference: small shapes that cluster together = artwork;
big simple shapes = layout.

USAGE
-----
  python3 pdf_to_editable_pptx.py input.pdf output.pptx
  python3 pdf_to_editable_pptx.py input.pdf output.pptx \
      --dpi 150 --preview-dir previews/ --skip-rect 3:410,110,490,340

  --dpi N           how sharp to render the art layer (150 is plenty)
  --preview-dir D   also save each page's art layer as a PNG so you can
                    see exactly what will go in the background
  --skip-rect P:x0,y0,x1,y1
                    ignore artwork inside this rectangle on page P
                    (coordinates are PDF "points"; 72 points = 1 inch;
                    measured from the TOP-LEFT corner of the page).
                    Repeat the flag for more rectangles.
  --remove-text "PHRASE"
                    also remove every text line containing PHRASE
                    (case doesn't matter). Repeat for more phrases.
                    The default list already removes the download
                    watermark and the "Copying is not permitted" footer.

KNOWN LIMITS (worth knowing before you judge the output)
--------------------------------------------------------
  * Fonts: PowerPoint substitutes any font you don't have installed, so
    text may look slightly wider or narrower than the PDF.
  * Tables are rebuilt as separate boxes, lines, and text - not as native
    PowerPoint tables - but every piece is still editable.

WHAT YOU NEED INSTALLED
-----------------------
  pip install PyMuPDF python-pptx pillow numpy
"""

# ============================================================================
# IMPORTS - bringing in the toolboxes this script needs
# ============================================================================
# "import" loads code other people wrote so we don't have to write it
# ourselves. Each library below is a specialized toolbox.

import argparse   # reads the options you type after the script name
import io         # lets us hold a file (like a PNG) in memory, not on disk
import sys        # lets us print warnings to the error channel

import fitz       # PyMuPDF: opens PDFs to read their text, shapes, images
import numpy as np                     # fast math on grids of pixel numbers
from PIL import Image, ImageDraw       # Pillow: builds and edits images
from pptx import Presentation          # python-pptx: builds PowerPoint files
from pptx.util import Emu, Pt          # PowerPoint's measurement units
from pptx.dml.color import RGBColor    # PowerPoint's way of writing colors
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR  # shape catalogs
from pptx.enum.text import MSO_AUTO_SIZE               # text box size rules
from pptx.oxml import parse_xml        # lets us write raw PowerPoint XML
from pptx.oxml.ns import qn            # helper for XML tag names

# One inch equals 72 PDF "points" and 914,400 "EMUs" (PowerPoint's tiny
# internal unit). These constants convert between the two systems.
POINTS_PER_INCH = 72.0
EMU_PER_POINT = 914400 / 72  # = 12,700 EMUs in one point

# Some PDFs typeset math symbols with the "Mathematical Pi" font. In that
# font, the characters stored in the file are NOT what you see on paper:
# the file stores "5" but the printed symbol is "=". This little dictionary
# translates the stored character (left) into the real symbol (right).
MATH_PI_SYMBOLS = {"1": "+", "2": "\u2212", "3": "\u00D7",   # + minus times
                   "4": "\u00F7", "5": "="}                  # divide equals

# Any line of text containing one of these phrases (upper/lower case does
# not matter) is treated as a watermark and removed from the output. You
# can add more phrases on the command line with --remove-text.
DEFAULT_REMOVE_PHRASES = [
    "downloaded by",                 # the per-user download watermark
    "copying is not permitted",      # the publisher's footer notice
]


# ============================================================================
# SMALL HELPERS - tiny functions used all over the script
# ============================================================================

def pt_to_emu(points):
    """Convert a PDF measurement (points) into PowerPoint's EMU unit.

    A 'def' line creates a reusable mini-program (a "function"); whatever
    follows 'return' is the answer it hands back when you call it.
    """
    return Emu(int(points * EMU_PER_POINT))


def float_color_to_rgb(color_tuple):
    """Turn a PDF color like (1.0, 0.85, 0.2) into a PowerPoint RGBColor.

    PDFs store colors as three decimals from 0 to 1 (red, green, blue).
    PowerPoint wants whole numbers from 0 to 255, so we multiply by 255.
    """
    if color_tuple is None:
        return None
    r, g, b = (max(0, min(255, round(c * 255))) for c in color_tuple[:3])
    return RGBColor(r, g, b)


def int_color_to_rgb(color_int):
    """Turn a PDF text color (one packed number) into a PowerPoint RGBColor.

    PDF text colors arrive packed into a single integer. The '>>' and '&'
    operations slice that integer into its red, green, and blue bytes -
    like cutting one long number into three two-digit pieces.
    """
    return RGBColor((color_int >> 16) & 255, (color_int >> 8) & 255,
                    color_int & 255)


def rect_overlap_fraction(inner, outer):
    """How much of rectangle 'inner' sits inside rectangle 'outer'? (0 to 1)

    Used for questions like "is this speech bubble mostly inside the
    character's art cluster?" A result of 1.0 means fully inside.
    """
    ix = fitz.Rect(inner) & fitz.Rect(outer)   # '&' = the overlapping part
    if ix.is_empty:
        return 0.0
    inner_area = inner.width * inner.height
    if inner_area <= 0:
        return 0.0
    return (ix.width * ix.height) / inner_area


def get_visible_drawings(page):
    """Return the page's vector shapes - but ONLY the visible ones.

    Here's a sneaky thing PDFs do: an artist might draw a WHOLE character
    and then use a "clip" (an invisible stencil window) so only the top
    half actually shows - the rest is silently cropped away. The simple
    page.get_drawings() call reports those hidden parts as if they were
    on the page, which would confuse every decision this script makes.

    So we ask for the extended version, which also reports each clip
    window ("scissor") and how the shapes nest inside them ("level",
    like indentation levels in an outline). As we walk the list we keep
    a stack (a pile) of the clip windows currently in force; any shape
    whose box falls at least 90% outside its clip windows is invisible
    on the real page, so we leave it out.
    """
    visible = []
    clip_stack = []                       # the pile of active clip windows
    for d in page.get_drawings(extended=True):
        level = d.get("level", 0)
        # Leaving a nesting level means its clip windows no longer apply:
        # pop (remove) every clip that is at this level or deeper.
        while clip_stack and clip_stack[-1][0] >= level:
            clip_stack.pop()
        if d["type"] == "clip":
            scissor = d.get("scissor")
            if scissor is not None:
                clip_stack.append((level, fitz.Rect(scissor)))
            continue                      # a clip is a window, not a shape
        if d["type"] == "group":
            continue                      # grouping info, not a shape
        r = fitz.Rect(d["rect"])
        if r.is_empty:
            visible.append(d)
            continue
        # Shrink the shape's box by every active clip window. Whatever
        # remains is the part that actually shows on the page.
        shown = fitz.Rect(r)
        for _level, scissor in clip_stack:
            shown = shown & scissor
        area = r.width * r.height
        shown_area = (0 if shown.is_empty
                      else shown.width * shown.height)
        if area > 0 and shown_area / area < 0.30:
            continue                      # 70%+ cropped away = treat as
                                          # invisible (a sliver isn't worth
                                          # rebuilding at full size)
        visible.append(d)
    return visible


# ============================================================================
# STEP 1: FIND THE ARTWORK ("seeds" -> "clusters")
# ============================================================================
# A "seed" is one small piece of probable artwork: a small curved vector
# shape or a small embedded picture. Characters are made of dozens of
# seeds packed together; we group nearby seeds into "clusters".

def collect_seeds(page):
    """Gather every small piece of probable artwork on one PDF page."""
    seeds = []

    # --- Part A: small CURVED vector shapes -------------------------------
    # get_visible_drawings() returns every visible vector shape. Each has
    # (the box it fits inside) and "items" - the individual strokes:
    # "l" = straight line, "re" = rectangle, "c" = curve. Artwork is full
    # of curves; layout mostly isn't.
    for d in get_visible_drawings(page):
        r = fitz.Rect(d["rect"])
        if r.is_empty:
            continue                       # skip shapes with no size
        if not any(item[0] == "c" for item in d["items"]):
            continue                       # no curves -> layout, not art
        # See-through shapes (under 90% opacity) are soft drop shadows.
        # Their boxes sprawl over tables and text and would form phantom
        # clusters that drag junk into the art layer - skip them.
        if (d.get("fill_opacity") or 1) < 0.9 and d.get("color") is None:
            continue
        # Big curved shapes (like a giant rounded text box) are layout,
        # not art, so only small ones count as seeds.
        if r.width * r.height > 6000 or r.width > 150 or r.height > 150:
            continue
        seeds.append(r)

    # --- Part B: small COLOR pictures --------------------------------------
    page_area = page.rect.width * page.rect.height
    for info in page.get_image_info():
        # Gray-only images in workbooks are almost always soft drop
        # shadows. Their boxes stretch across boxes and tables and would
        # drag junk into the art layer, so skip them.
        if "gray" in (info.get("cs-name") or "").lower():
            continue
        r = fitz.Rect(info["bbox"])
        if r.is_empty:
            continue
        coverage = (r.width * r.height) / page_area
        # Pictures that blanket the page are background textures (handled
        # separately in Step 2); near-zero ones are junk.
        if coverage > 0.25 or coverage < 0.0001:
            continue
        if r.width > 200 or r.height > 250:
            continue
        seeds.append(r)

    return seeds


def drop_seeds_on_text(page, seeds, max_overlap=0.2):
    """Throw away seeds that sit mostly on top of body text.

    Some shapes (like a soft glow under a heading) overlap real sentences.
    Keeping them would glue chunks of body text into the art layer. So:
    if more than 20% of a seed's box is covered by words, drop the seed.
    """
    # page.get_text("words") gives every word with its rectangle.
    words = [fitz.Rect(w[:4]) for w in page.get_text("words")]
    survivors = []
    for s in seeds:
        seed_area = s.width * s.height
        if seed_area <= 0:
            continue
        covered = 0.0
        for w in words:
            ix = fitz.Rect(s) & w        # the part of this word inside the seed
            if not ix.is_empty:
                covered += ix.width * ix.height
        if covered / seed_area <= max_overlap:
            survivors.append(s)
    return survivors


def merge_into_clusters(seed_boxes, pad=4):
    """Group seeds that sit near each other into bigger 'cluster' boxes.

    Imagine each seed box growing 4 points fatter on every side; any boxes
    that now touch get glued into one. Repeat until nothing new touches.
    A cartoon character's dozens of seeds end up as ONE cluster box.
    """
    boxes = [fitz.Rect(b) for b in seed_boxes]   # work on copies
    changed = True
    while changed:            # keep looping until a full pass glues nothing
        changed = False
        merged = []
        while boxes:
            b = boxes.pop()                                  # take one box
            grown = fitz.Rect(b) + (-pad, -pad, pad, pad)    # fatten it
            i = 0
            while i < len(boxes):
                if grown.intersects(boxes[i]):     # does it touch another?
                    b = fitz.Rect(b) | boxes[i]    # '|' = box around both
                    grown = fitz.Rect(b) + (-pad, -pad, pad, pad)
                    boxes.pop(i)                   # absorbed; remove it
                    changed = True
                else:
                    i += 1
            merged.append(b)
        boxes = merged
    return boxes


def keep_dense_clusters(clusters, seeds, min_seeds=2, max_area_in2=10.0):
    """Keep only clusters that look like real artwork.

    Real artwork = several seeds packed into a reasonably small box. One
    lonely squiggle, or a 'cluster' the size of half the page, is probably
    noise. Returns a list of (cluster_box, [member seed boxes]) pairs.
    """
    kept = []
    for c in clusters:
        members = [s for s in seeds if c.intersects(s)]
        area_in2 = (c.width / 72.0) * (c.height / 72.0)   # square inches
        if (len(members) >= min_seeds and area_in2 <= max_area_in2
                and c.width > 10 and c.height > 10):
            kept.append((fitz.Rect(c), members))
    return kept


def adopt_attached_shapes(page, clusters):
    """Pull bigger shapes that belong to an art cluster into that cluster.

    Example: a speech bubble. The bubble outline is one LARGE curved shape,
    so the seed step skipped it - but it clearly belongs with the character
    it points to. Rule: if at least 60% of a shape sits inside an existing
    cluster, "adopt" it (and stretch the cluster box to fit). We remember
    adopted shapes separately because any TEXT inside them (the bubble's
    words) should travel with the art instead of becoming editable.
    """
    page_area = page.rect.width * page.rect.height
    adopted = []
    for _pass in range(2):     # two passes: adopting one shape can pull the
        for d in get_visible_drawings(page):                 # next one into range
            r = fitz.Rect(d["rect"])
            if r.is_empty:
                continue
            if not any(item[0] == "c" for item in d["items"]):
                continue                              # only curved shapes
            if r.width * r.height > 0.15 * page_area:
                continue                              # page-sized = furniture
            for idx, (box, members) in enumerate(clusters):
                already = any(abs(m.x0 - r.x0) < 0.01 and
                              abs(m.y0 - r.y0) < 0.01 and
                              abs(m.x1 - r.x1) < 0.01 for m in members)
                if already:
                    break                             # joined earlier; skip
                if rect_overlap_fraction(r, box) >= 0.6:
                    members.append(fitz.Rect(r))            # join the family
                    clusters[idx] = (box | r, members)      # stretch the box
                    adopted.append(fitz.Rect(r))
                    break
    return clusters, adopted


def words_inside_adopted(page, adopted):
    """Find the words that live INSIDE adopted shapes (e.g. bubble text).

    These words are part of the artwork: they stay in the background
    picture and do NOT become editable text. A word only counts if at
    least 90% of its rectangle sits inside the adopted shape - a word
    merely brushing against a character's edge should stay editable.
    """
    art_words = []
    for w in page.get_text("words"):
        wrect = fitz.Rect(w[:4])
        for a in adopted:
            if rect_overlap_fraction(wrect, a) >= 0.9:
                art_words.append(wrect)
                break
    return art_words


def find_removed_text_lines(page, phrases):
    """Find the watermark lines that should disappear from the output.

    We read every line of text on the page, squash it to lower case
    (so "Downloaded" and "downloaded" match the same way), and check
    whether it contains any phrase from the removal list. Matching lines
    are returned as rectangles, which later steps use to also remove the
    little backing strip drawn behind the watermark.
    """
    removed = []
    if not phrases:
        return removed
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            text = "".join(s["text"] for s in line["spans"]).lower()
            if any(p in text for p in phrases):
                removed.append(fitz.Rect(line["bbox"]))
    return removed


def extract_texture_image(doc, xref, smask_xref):
    """Pull one embedded picture out of the PDF as a Pillow image.

    Big page textures (the colored page borders and header bands) are
    stored inside the PDF as pictures. doc.extract_image() hands us the
    raw picture bytes; Pillow turns those bytes into an image we can
    resize and paste. If the picture has an "smask" (a soft transparency
    stencil), we attach it as the image's alpha (see-through) channel so
    rounded corners stay rounded.
    """
    try:
        raw = doc.extract_image(xref)
        img = Image.open(io.BytesIO(raw["image"]))
        if img.mode == "CMYK":            # print-style colors -> screen colors
            img = img.convert("RGB")
        img = img.convert("RGBA")         # RGBA = colors + transparency
        if smask_xref:                    # apply the transparency stencil
            mraw = doc.extract_image(smask_xref)
            mask = Image.open(io.BytesIO(mraw["image"])).convert("L")
            if mask.size != img.size:
                mask = mask.resize(img.size)
            img.putalpha(mask)
        return img
    except Exception as exc:              # a damaged picture shouldn't crash us
        print(f"  (skipped one texture: {exc})", file=sys.stderr)
        return None


def find_content_panels(page):
    """Find big light-colored 'paper' panels drawn over the page texture.

    Workbook pages often paint a colorful texture across the whole page,
    then place a large near-white rounded panel on top to hold the lesson
    content. We need to know where those panels are so the background
    canvas shows clean paper there instead of raw texture.
    Returns a list of (rectangle, fill_color) pairs in painting order.
    """
    panels = []
    page_area = page.rect.width * page.rect.height
    for d in get_visible_drawings(page):
        fill = d.get("fill")
        if fill is None:
            continue
        r = fitz.Rect(d["rect"])
        big_enough = r.width * r.height > 0.40 * page_area
        brightness = sum(fill[:3]) / 3            # 1.0 = pure white
        if big_enough and brightness > 0.85:
            panels.append((r, float_color_to_rgb(fill)))
    return panels


def build_background_canvas(doc, page, canvas_size, scale):
    """Create the locked background picture: page textures + paper panels.

    This image goes into the slide's background, where nothing can bump
    it. It holds ONLY the page's decoration that nobody ever needs to
    move: the big texture pictures (colored borders, header bands) and
    the light "paper" panels painted over them. The artwork is NOT in
    here anymore - each art piece becomes its own movable picture.

    The plan, in order (later layers cover earlier ones):
      a. Start with a blank white canvas the size of the page.
      b. Paste the page's big TEXTURE pictures, scaled to where the page
         places them (their rectangles often hang off the page edges, so
         we keep only the on-page part).
      c. Paint the big light 'paper' panels back over the texture, so the
         content area is clean paper again.
    """
    # (a) the blank white canvas
    canvas = Image.new("RGB", canvas_size, (255, 255, 255))

    # (b) paste the big texture pictures
    page_area = page.rect.width * page.rect.height
    for info in page.get_image_info(xrefs=True):
        r = fitz.Rect(info["bbox"])
        if r.is_empty:
            continue
        coverage = (r.width * r.height) / page_area
        if coverage <= 0.25:
            continue                      # small pictures are art, not texture
        img = extract_texture_image(doc, info.get("xref", 0),
                                    info.get("smask", 0))
        if img is None or info.get("xref", 0) == 0:
            continue
        full_w = max(int(r.width * scale), 1)
        full_h = max(int(r.height * scale), 1)
        img = img.resize((full_w, full_h))
        visible = fitz.Rect(r) & page.rect          # the on-page portion
        if visible.is_empty:
            continue
        crop = img.crop((int((visible.x0 - r.x0) * scale),
                         int((visible.y0 - r.y0) * scale),
                         int((visible.x1 - r.x0) * scale),
                         int((visible.y1 - r.y0) * scale)))
        # The third argument makes Pillow respect the transparency channel.
        canvas.paste(crop, (int(visible.x0 * scale), int(visible.y0 * scale)),
                     crop)

    # (c) repaint the big light 'paper' panels over the texture
    draw = ImageDraw.Draw(canvas)
    for panel_rect, panel_color in find_content_panels(page):
        pr = fitz.Rect(panel_rect) & page.rect
        draw.rectangle([pr.x0 * scale, pr.y0 * scale,
                        pr.x1 * scale, pr.y1 * scale],
                       fill=(panel_color[0], panel_color[1], panel_color[2]))

    return canvas


def build_cluster_picture(page, photo, cluster_box, members, adopted,
                          art_words, removed_rects, scale):
    """Turn ONE art cluster into a movable picture with see-through edges.

    The picture is built like a sticker:
      a. Start with a fully TRANSPARENT canvas the size of the cluster.
      b. Copy each member piece's little rectangle from the page photo
         onto it (adopted shapes like speech bubbles go first, so the
         character lands on top of its bubble).
      c. Punch see-through holes over anything that rode along by
         accident: words that are not part of the art, thin answer/grid
         lines passing through, and any removed watermark text. Holes
         (instead of color patches) matter here because the picture can
         be MOVED - a painted patch would travel with it and look wrong.
      d. If almost nothing visible remains, report "nothing here" (None).

    Returns (picture, rectangle_on_page) or None.
    """
    # (a) clip the cluster to the page and make the transparent canvas.
    rect = fitz.Rect(cluster_box) & page.rect
    if rect.is_empty:
        return None
    box_px = (int(rect.x0 * scale), int(rect.y0 * scale),
              int(rect.x1 * scale), int(rect.y1 * scale))
    width_px = box_px[2] - box_px[0]
    height_px = box_px[3] - box_px[1]
    if width_px <= 2 or height_px <= 2:
        return None
    # "RGBA" = red, green, blue, and Alpha (transparency). Alpha 0 means
    # fully see-through, which is what the whole sticker starts as.
    sticker = Image.new("RGBA", (width_px, height_px), (0, 0, 0, 0))

    # (b) copy each member piece onto the sticker.
    adopted_keys = {(round(a.x0, 1), round(a.y0, 1)) for a in adopted}

    def paste_member(m):
        m = fitz.Rect(m) & rect                      # stay inside the sticker
        if m.is_empty:
            return
        crop_box = (int(m.x0 * scale), int(m.y0 * scale),
                    int(m.x1 * scale), int(m.y1 * scale))
        if crop_box[2] <= crop_box[0] or crop_box[3] <= crop_box[1]:
            return
        crop = photo.crop(crop_box)
        # Skip almost-blank pieces: in grayscale ("L"), count how many
        # pixels are darker than near-white. Under 5% means "blank".
        gray = np.asarray(crop.convert("L"))
        if (gray < 235).mean() < 0.05:
            return
        # Place the piece at its position RELATIVE to the sticker's
        # top-left corner (page position minus sticker position).
        sticker.paste(crop.convert("RGBA"),
                      (crop_box[0] - box_px[0], crop_box[1] - box_px[1]))

    bubbles_first = [m for m in members
                     if (round(m.x0, 1), round(m.y0, 1)) in adopted_keys]
    everyone_else = [m for m in members
                     if (round(m.x0, 1), round(m.y0, 1)) not in adopted_keys]
    for m in bubbles_first + everyone_else:
        paste_member(m)

    # (c) punch see-through holes over the accidental passengers.
    def punch_hole(r, pad=1.5):
        """Make a rectangle of the sticker fully transparent."""
        r = (fitz.Rect(r) + (-pad, -pad, pad, pad)) & rect
        if r.is_empty:
            return
        hole = ImageDraw.Draw(sticker)
        hole.rectangle([r.x0 * scale - box_px[0], r.y0 * scale - box_px[1],
                        r.x1 * scale - box_px[0], r.y1 * scale - box_px[1]],
                       fill=(0, 0, 0, 0))

    art_word_keys = {(round(w.x0, 1), round(w.y0, 1)) for w in art_words}
    for w in page.get_text("words"):
        wrect = fitz.Rect(w[:4])
        if not wrect.intersects(rect):
            continue                                  # nowhere near us
        if (round(wrect.x0, 1), round(wrect.y0, 1)) in art_word_keys:
            continue                                  # bubble text stays!
        punch_hole(wrect)

    for rr in removed_rects:                          # watermark lines
        if fitz.Rect(rr).intersects(rect):
            punch_hole(rr, pad=3)

    # Thin straight rules (answer lines, table grid lines) that cut
    # through the sticker - unless the line belongs to this art.
    for d in get_visible_drawings(page):
        r = fitz.Rect(d["rect"])
        if not r.intersects(rect):
            continue
        if any(item[0] == "c" for item in d["items"]):
            continue                                  # curved = art; leave it
        w = max(r.width, 0.1)
        h = max(r.height, 0.1)
        if max(w / h, h / w) <= 8:
            continue                                  # not thin -> not a rule
        center = fitz.Point((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2)
        if rect.contains(center):
            continue                                  # belongs to art; leave it
        punch_hole(r)

    # (d) is anything meaningful left? Look only at pixels that are both
    # visible (alpha above 0) and darker than near-white.
    arr = np.asarray(sticker)
    visible_ink = (arr[:, :, 3] > 0) & (arr[:, :, :3].mean(axis=2) < 235)
    if visible_ink.mean() < 0.01:
        return None

    return sticker, rect


# ============================================================================
# STEP 3: PUT THE TEXTURES INTO THE SLIDE BACKGROUND (non-editable!)
# ============================================================================

# PowerPoint files are secretly folders of XML text. python-pptx covers
# most jobs, but "use this picture as the slide background" isn't built
# in, so we write that one piece of XML ourselves. These lines declare the
# XML dialects (namespaces) PowerPoint expects to see.
XML_NAMESPACES = (
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)


def set_slide_background(slide, png_bytes):
    """Make a PNG the slide's background. Backgrounds can't be edited.

    How: register the picture inside the PowerPoint file (which hands back
    an ID), then insert a <p:bg> block at the start of the slide's XML
    that says "fill the background by stretching picture <ID>".
    """
    image_part, rid = slide.part.get_or_add_image_part(io.BytesIO(png_bytes))
    cSld = slide._element.find(qn("p:cSld"))        # the slide's XML root
    old_bg = cSld.find(qn("p:bg"))
    if old_bg is not None:
        cSld.remove(old_bg)                         # replace any old one
    bg = parse_xml(
        f'<p:bg {XML_NAMESPACES}><p:bgPr><a:blipFill>'
        f'<a:blip r:embed="{rid}"/>'
        f'<a:stretch><a:fillRect/></a:stretch>'
        f'</a:blipFill><a:effectLst/></p:bgPr></p:bg>'
    )
    cSld.insert(0, bg)                  # backgrounds must come first in XML


# ============================================================================
# STEP 4: REBUILD THE LAYOUT AS EDITABLE POWERPOINT SHAPES
# ============================================================================

def mostly_inside_clusters(rect, cluster_list, threshold=0.4):
    """Is a big chunk of 'rect' inside any art cluster? (True/False)

    Used to skip recreating shapes that are really character body parts -
    the background picture already shows them. The bar is 40%: character
    pieces poke partway out of their cluster box, while genuine layout
    boxes a character merely leans over stay well under this.
    """
    return any(rect_overlap_fraction(rect, box) >= threshold
               for box, _ in cluster_list)


def add_layout_shapes(slide, page, clusters, adopted, removed_rects=()):
    """Recreate boxes, bands, table fills, and lines as editable shapes.

    We walk the page's vector drawings IN ORDER (PDFs paint from back to
    front, and so does PowerPoint, so keeping the order keeps the stacking
    right) and rebuild each one:
      * filled shapes  -> rectangles (rounded if the original had curves)
      * stroked lines  -> straight connector lines
    Anything that belongs to the artwork is skipped - the movable art
    pictures already show it. Anything sitting behind a removed watermark
    line (its little backing strip) is skipped too.
    """
    page_area = page.rect.width * page.rect.height
    adopted_keys = {(round(a.x0, 1), round(a.y0, 1)) for a in adopted}
    drawings = get_visible_drawings(page)

    # Note every solid (opaque) filled shape and WHEN it was painted, so
    # we can ask: "is this shape later painted over completely?" PDFs
    # sometimes use solid black rectangles as internal stencils that end
    # up fully buried under the artwork - if we rebuilt those, they would
    # show up as ugly black slabs.
    opaque_fills = []
    for index, d in enumerate(drawings):
        if d.get("fill") is None:
            continue
        if (d.get("fill_opacity") or 1) < 0.9:
            continue
        fr = fitz.Rect(d["rect"])
        if fr.is_empty or fr.width * fr.height > 0.40 * page_area:
            continue                # giant paper panels don't count as cover
        opaque_fills.append((index, fr))

    def covered_by_later_fills(my_index, rect, samples=5, needed=0.7):
        """Is 'rect' at least 70% hidden under fills painted after it?

        We check a 5-by-5 grid of test points spread across the
        rectangle. For each point we ask: does any LATER opaque fill sit
        on top of it? If 70% or more of the points are covered, the shape
        is effectively invisible on the real page, so the caller skips it.
        """
        later = [fr for j, fr in opaque_fills if j > my_index]
        if not later:
            return False
        hits = 0
        for row in range(samples):
            for col in range(samples):
                # Walk the grid: (col + 0.5) / samples gives evenly spaced
                # positions like 0.1, 0.3, 0.5, 0.7, 0.9 across the width.
                point = fitz.Point(
                    rect.x0 + (col + 0.5) / samples * rect.width,
                    rect.y0 + (row + 0.5) / samples * rect.height)
                if any(fr.contains(point) for fr in later):
                    hits += 1
        return hits / (samples * samples) >= needed

    for index, d in enumerate(drawings):
        r = fitz.Rect(d["rect"])
        # Skip only true specks. Careful: a perfectly horizontal line has
        # ZERO height (its box is "empty" in math terms) but it's a real
        # answer line we must keep - so we test both directions and only
        # drop shapes tiny in BOTH.
        if r.width < 0.5 and r.height < 0.5:
            continue

        has_curves = any(item[0] == "c" for item in d["items"])
        has_rect_item = any(item[0] == "re" for item in d["items"])

        # ---- skip things that belong to the art layer --------------------
        if (round(r.x0, 1), round(r.y0, 1)) in adopted_keys:
            continue                          # e.g. the speech bubble itself
        if mostly_inside_clusters(r, clusters):
            continue                          # character pieces, icon parts

        # ---- skip the strip behind removed watermark text ------------------
        # Watermark lines often sit on a little colored strip. The text is
        # gone, so the strip should go too. (Grow each removed rectangle a
        # few points so the strip's slightly larger edges still match.)
        if removed_rects and any(
                rect_overlap_fraction(r, fitz.Rect(rr) + (-4, -4, 4, 4)) >= 0.6
                for rr in removed_rects):
            continue

        # ---- skip see-through "shadow" shapes -----------------------------
        # Drop shadows are painted at 20-40% opacity in the PDF. PowerPoint
        # shapes here would come out fully solid (a black slab), so we
        # leave shadows out entirely - the page looks cleaner anyway.
        if (d.get("fill_opacity") or 1) < 0.9:
            continue

        fill = float_color_to_rgb(d.get("fill"))
        stroke = float_color_to_rgb(d.get("color"))
        stroke_width = d.get("width") or 0

        # ---- CASE 1: a filled area  ->  a rectangle shape ------------------
        if fill is not None:
            # A filled shape covering nearly half the page is the 'paper'
            # panel. The background canvas already painted it (Step 2c),
            # and a giant shape here would hide the artwork - skip it.
            if r.width * r.height > 0.40 * page_area:
                continue
            # A shape buried under later fills is invisible on the real
            # page (often an internal stencil rectangle) - skip it too.
            if covered_by_later_fills(index, r):
                continue
            # Small filled doodles with neither straight rectangles nor
            # curves in them (star shapes, decorative polygons) would turn
            # into ugly colored slabs - they belong to decoration, skip.
            if not has_rect_item and not has_curves and \
                    r.width * r.height < 3000:
                continue
            shape_kind = (MSO_SHAPE.ROUNDED_RECTANGLE if has_curves
                          else MSO_SHAPE.RECTANGLE)
            shp = slide.shapes.add_shape(
                shape_kind,
                pt_to_emu(r.x0), pt_to_emu(r.y0),
                pt_to_emu(max(r.width, 1)), pt_to_emu(max(r.height, 1)))
            shp.fill.solid()
            shp.fill.fore_color.rgb = fill
            if stroke is not None and stroke_width > 0:
                shp.line.color.rgb = stroke
                shp.line.width = Pt(max(stroke_width, 0.25))
            else:
                shp.line.fill.background()            # no outline
            shp.shadow.inherit = False                # no surprise shadows

        # ---- CASE 2: stroke-only  ->  straight connector lines -------------
        elif stroke is not None:
            for item in d["items"]:
                if item[0] == "l":                    # one line segment
                    p1, p2 = item[1], item[2]
                    ln = slide.shapes.add_connector(
                        MSO_CONNECTOR.STRAIGHT,
                        pt_to_emu(p1.x), pt_to_emu(p1.y),
                        pt_to_emu(p2.x), pt_to_emu(p2.y))
                    ln.line.color.rgb = stroke
                    ln.line.width = Pt(max(stroke_width, 0.5))
                elif item[0] == "re":                 # an outlined rectangle
                    rr = fitz.Rect(item[1])
                    shp = slide.shapes.add_shape(
                        MSO_SHAPE.RECTANGLE,
                        pt_to_emu(rr.x0), pt_to_emu(rr.y0),
                        pt_to_emu(max(rr.width, 1)),
                        pt_to_emu(max(rr.height, 1)))
                    shp.fill.background()             # see-through inside
                    shp.line.color.rgb = stroke
                    shp.line.width = Pt(max(stroke_width, 0.5))
                    shp.shadow.inherit = False


def add_editable_text(slide, page, art_words, remove_phrases=()):
    """Recreate every line of text as an editable text box.

    page.get_text("dict") hands back the text organized as:
        blocks -> lines -> spans
    A "span" is a run of letters sharing one style (font, size, color).
    We make one text box per LINE, and inside it one styled "run" per
    span - so a line like "The 3 has a value of 30" keeps its red 3.
    Words that belong to artwork (speech-bubble text) are skipped here
    because they already live in a movable art picture. Lines containing
    a phrase from the removal list (the watermark footers) are skipped
    entirely.
    """
    art_word_keys = {(round(w.x0, 1), round(w.y0, 1)) for w in art_words}

    text_data = page.get_text("dict")
    for block in text_data["blocks"]:
        if block.get("type") != 0:      # type 0 = text (1 would be an image)
            continue
        for line in block["lines"]:
            if line.get("dir", (1, 0))[0] < 0.9:
                continue                # skip rotated/vertical text (rare)
            lrect = fitz.Rect(line["bbox"])
            spans = [s for s in line["spans"] if s["text"].strip()]
            if not spans:
                continue

            # Watermark check: squash the whole line to lower case and
            # see if it contains any phrase from the removal list.
            line_text = "".join(s["text"] for s in spans).lower()
            if any(p in line_text for p in remove_phrases):
                continue                  # leave the watermark behind

            # If this line's words belong to artwork, leave the whole line
            # to the background picture.
            span_keys = {(round(fitz.Rect(s["bbox"]).x0, 1),
                          round(fitz.Rect(s["bbox"]).y0, 1)) for s in spans}
            if span_keys & art_word_keys or \
                    (round(lrect.x0, 1), round(lrect.y0, 1)) in art_word_keys:
                continue

            # Make the text box, with a little breathing room on the right
            # and bottom so PowerPoint never clips the letters.
            box = slide.shapes.add_textbox(
                pt_to_emu(lrect.x0), pt_to_emu(lrect.y0),
                pt_to_emu(lrect.width * 1.10 + 4),
                pt_to_emu(lrect.height * 1.25 + 2))
            # Text boxes can inherit a solid white fill from the theme,
            # which would hide the artwork behind them. Force the box
            # itself to be completely see-through (only letters show).
            box.fill.background()
            box.line.fill.background()
            tf = box.text_frame
            tf.word_wrap = False                    # one line stays one line
            tf.auto_size = MSO_AUTO_SIZE.NONE       # don't resize on me
            # PowerPoint pads text boxes by default; zero the padding so our
            # letters land exactly where the PDF put them.
            tf.margin_left = tf.margin_right = 0
            tf.margin_top = tf.margin_bottom = 0

            paragraph = tf.paragraphs[0]            # the (only) paragraph
            for span in spans:
                text = span["text"]
                font_name = span["font"]
                # Mathematical Pi fonts store symbols as digits (see the
                # MATH_PI_SYMBOLS note near the top). Translate each
                # character so "5" becomes "=" and "3" becomes "x".
                if "mathematicalpi" in font_name.lower():
                    text = "".join(MATH_PI_SYMBOLS.get(ch, ch) for ch in text)

                run = paragraph.add_run()           # one styled chunk
                # PDFs sometimes hide invisible "control characters" in
                # text (leftovers from list markers). PowerPoint would
                # print them as odd codes like _x0007_, so strip out any
                # character below the regular printable range.
                run.text = "".join(ch for ch in text if ord(ch) >= 32)
                run.font.size = Pt(max(span["size"], 1))
                run.font.color.rgb = int_color_to_rgb(span["color"])
                # The span's "flags" number packs style switches into its
                # binary digits ("bits"). The '&' check reads one bit:
                # bit 16 = bold, bit 2 = italic.
                run.font.bold = bool(span["flags"] & 16)
                run.font.italic = bool(span["flags"] & 2)
                # Font names sometimes carry a random prefix like
                # "ABCDEE+Arial-Bold"; the real family name is the part
                # after '+' and before '-'.
                clean = font_name.split("+")[-1].split("-")[0]
                run.font.name = clean or "Arial"


# ============================================================================
# STEP 5: THE MAIN PROGRAM - tie all the steps together
# ============================================================================

def parse_skip_rect(spec):
    """Decode a --skip-rect value like '3:410,110,490,340'."""
    try:
        page_str, coords = spec.split(":")
        x0, y0, x1, y1 = (float(v) for v in coords.split(","))
        return int(page_str) - 1, fitz.Rect(x0, y0, x1, y1)
    except Exception:
        raise argparse.ArgumentTypeError(
            f"bad --skip-rect '{spec}'; expected PAGE:x0,y0,x1,y1")


def main():
    # ---- read the command-line options -------------------------------------
    ap = argparse.ArgumentParser(
        description="Convert a PDF into an editable PPTX with the decorative "
                    "artwork locked into each slide's background.",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pdf", help="the source PDF file")
    ap.add_argument("pptx_out", help="where to save the PowerPoint file")
    ap.add_argument("--dpi", type=int, default=150,
                    help="art layer sharpness (default 150)")
    ap.add_argument("--skip-rect", type=parse_skip_rect, action="append",
                    default=[], metavar="PAGE:x0,y0,x1,y1",
                    help="ignore artwork in this region (repeatable)")
    ap.add_argument("--preview-dir", default=None,
                    help="folder to save each page's background and art "
                         "pictures as PNGs")
    ap.add_argument("--remove-text", action="append", default=[],
                    metavar="PHRASE",
                    help="also remove text lines containing PHRASE "
                         "(repeatable; case-insensitive)")
    args = ap.parse_args()

    # The removal list = the built-in defaults + anything you added.
    # Everything is lower-cased once here so comparisons are simple.
    remove_phrases = [p.lower() for p in
                      DEFAULT_REMOVE_PHRASES + list(args.remove_text)]

    # ---- open the PDF and start a fresh PowerPoint --------------------------
    doc = fitz.open(args.pdf)
    prs = Presentation()
    # Size the slides to match the first PDF page exactly, so every
    # coordinate carries over one-to-one.
    first = doc[0].rect
    prs.slide_width = pt_to_emu(first.width)
    prs.slide_height = pt_to_emu(first.height)
    blank_layout = prs.slide_layouts[6]      # layout #6 is the empty one

    # Organize any --skip-rect values by page number for quick lookup.
    skips = {}
    for page_index, rect in args.skip_rect:
        skips.setdefault(page_index, []).append(rect)

    if args.preview_dir:
        import os
        os.makedirs(args.preview_dir, exist_ok=True)

    # ---- convert every page, one at a time ----------------------------------
    scale = args.dpi / POINTS_PER_INCH    # converts PDF points -> pixels
    for page_index in range(len(doc)):
        page = doc[page_index]
        slide = prs.slides.add_slide(blank_layout)

        # Photograph the real page once; the art stickers are cut from it.
        pix = page.get_pixmap(dpi=args.dpi)
        photo = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

        # STEP 0: find the watermark lines that must disappear.
        removed_rects = find_removed_text_lines(page, remove_phrases)

        # STEP 1: find the artwork.
        seeds = drop_seeds_on_text(page, collect_seeds(page))
        seeds = [s for s in seeds
                 if not any(s.intersects(k)
                            for k in skips.get(page_index, []))]
        clusters = keep_dense_clusters(merge_into_clusters(seeds), seeds)
        clusters, adopted = adopt_attached_shapes(page, clusters)
        art_words = words_inside_adopted(page, adopted)

        # STEP 2: build the locked background (textures + paper panels)
        # and put it into the slide. Nothing here ever needs moving.
        background = build_background_canvas(doc, page, photo.size, scale)
        if args.preview_dir:
            background.save(f"{args.preview_dir.rstrip('/')}/"
                            f"background-page-{page_index + 1}.png")
        png = io.BytesIO()
        background.save(png, format="PNG")
        set_slide_background(slide, png.getvalue())

        # STEP 3: rebuild boxes, bands, table grids, and lines as
        # editable shapes (the layer just above the background).
        add_layout_shapes(slide, page, clusters, adopted, removed_rects)

        # STEP 4: place each art cluster as its own MOVABLE picture.
        # Adding them after the shapes but before the text puts them
        # above the boxes (a character can lean over a box) and below
        # the words (text always stays readable).
        art_count = 0
        for box, members in clusters:
            result = build_cluster_picture(page, photo, box, members,
                                           adopted, art_words,
                                           removed_rects, scale)
            if result is None:
                continue                  # this cluster was empty/blank
            sticker, rect = result
            png = io.BytesIO()
            sticker.save(png, format="PNG")
            png.seek(0)                   # rewind so PowerPoint reads from
            slide.shapes.add_picture(     # the start of the picture data
                png,
                pt_to_emu(rect.x0), pt_to_emu(rect.y0),
                pt_to_emu(rect.width), pt_to_emu(rect.height))
            art_count += 1
            if args.preview_dir:
                sticker.save(f"{args.preview_dir.rstrip('/')}/"
                             f"art-page-{page_index + 1}-"
                             f"piece-{art_count}.png")

        # STEP 5: rebuild every line of text as editable text boxes,
        # skipping art words and the removed watermark lines.
        add_editable_text(slide, page, art_words, remove_phrases)

        print(f"page {page_index + 1}: {art_count} movable art pictures, "
              f"{len(art_words)} words kept with the art, "
              f"{len(removed_rects)} watermark lines removed")

    # ---- save and finish ------------------------------------------------------
    prs.save(args.pptx_out)
    print(f"saved {args.pptx_out}")


# This last bit means: "only run main() when this file is executed
# directly" (as opposed to being imported by another script as a toolbox).
if __name__ == "__main__":
    main()
