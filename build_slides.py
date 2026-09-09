from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

LOGO = '/Users/avanti.srihari/Documents/Claude/Kids-Portal/pbjtp_logo.png'
OUT  = '/Users/avanti.srihari/Documents/Claude/Kids-Portal/PBJTP_Kids_Slides.pptx'

# ── Brand colours ──────────────────────────────────────────────────────────────
MAROON    = RGBColor(0x7a, 0x1e, 0x32)
GREEN     = RGBColor(0x2d, 0x5a, 0x27)
GOLD      = RGBColor(0xC8, 0x94, 0x1A)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xFD, 0xF8, 0xF2)
LIGHT_MAROON = RGBColor(0xF5, 0xEB, 0xE8)
LIGHT_GREEN  = RGBColor(0xE8, 0xF5, 0xE4)
DARK      = RGBColor(0x1A, 0x1A, 0x1A)

W = Inches(13.33)   # widescreen width
H = Inches(7.5)     # widescreen height

prs = Presentation()
prs.slide_width  = W
prs.slide_height = H

blank_layout = prs.slide_layouts[6]   # completely blank

# ── Helpers ────────────────────────────────────────────────────────────────────

def add_rect(slide, x, y, w, h, fill, alpha=None):
    shape = slide.shapes.add_shape(1, x, y, w, h)
    shape.line.fill.background()
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    return shape

def add_text_box(slide, text, x, y, w, h,
                 font_size=24, bold=False, color=DARK,
                 align=PP_ALIGN.LEFT, font_name='Arial Rounded MT Bold',
                 wrap=True):
    txb = slide.shapes.add_textbox(x, y, w, h)
    txb.word_wrap = wrap
    tf  = txb.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.name = font_name
    run.font.color.rgb = color
    return txb

def add_bullet_box(slide, lines, x, y, w, h,
                   font_size=21, color=DARK, font_name='Arial',
                   line_color=None, bullet_char='●'):
    txb = slide.shapes.add_textbox(x, y, w, h)
    txb.word_wrap = True
    tf  = txb.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        if first:
            p = tf.paragraphs[0]; first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(4)
        run = p.add_run()
        run.text = f'{bullet_char}  {line}'
        run.font.size = Pt(font_size)
        run.font.name = font_name
        run.font.color.rgb = line_color or color
    return txb

def add_logo(slide, size=Inches(0.9)):
    slide.shapes.add_picture(LOGO,
        W - size - Inches(0.18),
        Inches(0.12), size, size)

def add_footer(slide, text='Pete Brown Jr. Tennis Program — Building Champions On and Off the Court'):
    rect = add_rect(slide, 0, H - Inches(0.42), W, Inches(0.42), MAROON)
    add_text_box(slide, text,
        Inches(0.2), H - Inches(0.40), W - Inches(0.4), Inches(0.38),
        font_size=12, color=WHITE, align=PP_ALIGN.CENTER,
        font_name='Arial', bold=False)

def colored_pill(slide, label, x, y, w, h, bg, text_color=WHITE, font_size=18):
    add_rect(slide, x, y, w, h, bg)
    add_text_box(slide, label, x, y + Inches(0.04), w, h - Inches(0.04),
                 font_size=font_size, bold=True, color=text_color,
                 align=PP_ALIGN.CENTER)

# ── TITLE SLIDE template ───────────────────────────────────────────────────────

def title_slide(title, subtitle, tagline, deck_num, deck_color):
    slide = prs.slides.add_slide(blank_layout)
    # full background
    add_rect(slide, 0, 0, W, H, OFF_WHITE)
    # left maroon stripe
    add_rect(slide, 0, 0, Inches(0.35), H, MAROON)
    # top band
    add_rect(slide, 0, 0, W, Inches(1.1), deck_color)
    # deck label
    add_text_box(slide, f'DECK {deck_num}', Inches(0.5), Inches(0.18),
                 Inches(3), Inches(0.7), font_size=22, bold=True, color=WHITE)
    # big logo centre-left
    slide.shapes.add_picture(LOGO, Inches(0.7), Inches(1.4), Inches(2.6), Inches(2.6))
    # title
    add_text_box(slide, title, Inches(3.8), Inches(1.3), Inches(8.8), Inches(2.0),
                 font_size=46, bold=True, color=MAROON, align=PP_ALIGN.LEFT)
    # subtitle
    add_text_box(slide, subtitle, Inches(3.8), Inches(3.3), Inches(8.8), Inches(1.0),
                 font_size=26, bold=False, color=GREEN, align=PP_ALIGN.LEFT)
    # tagline pill
    add_rect(slide, Inches(3.8), Inches(4.5), Inches(8.8), Inches(0.65), deck_color)
    add_text_box(slide, tagline, Inches(3.9), Inches(4.5), Inches(8.6), Inches(0.65),
                 font_size=19, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    # bottom accent
    add_rect(slide, 0, H - Inches(0.5), W, Inches(0.5), MAROON)
    add_text_box(slide, 'Pete Brown Jr. Tennis Program',
                 Inches(0.3), H - Inches(0.48), W - Inches(0.6), Inches(0.46),
                 font_size=14, color=WHITE, align=PP_ALIGN.CENTER)
    return slide

# ── CONTENT SLIDE template ─────────────────────────────────────────────────────

def content_slide(title, subtitle, bullets, accent_label=None, accent_text=None,
                  deck_color=MAROON, bg=OFF_WHITE, slide_num=None, total=None):
    slide = prs.slides.add_slide(blank_layout)
    add_rect(slide, 0, 0, W, H, bg)
    # top band
    add_rect(slide, 0, 0, W, Inches(1.35), deck_color)
    # left stripe
    add_rect(slide, 0, 0, Inches(0.22), H, deck_color)
    # title
    add_text_box(slide, title, Inches(0.45), Inches(0.12),
                 Inches(11.2), Inches(0.80), font_size=38, bold=True,
                 color=WHITE, align=PP_ALIGN.LEFT)
    # subtitle
    if subtitle:
        add_text_box(slide, subtitle, Inches(0.45), Inches(0.88),
                     Inches(11.2), Inches(0.42), font_size=20, bold=False,
                     color=WHITE, align=PP_ALIGN.LEFT, font_name='Arial')
    add_logo(slide)

    # body area
    body_y = Inches(1.5)
    body_h = Inches(5.3)

    if accent_label and accent_text:
        # two-column layout
        left_w = Inches(7.2)
        add_bullet_box(slide, bullets, Inches(0.45), body_y, left_w, body_h,
                       font_size=20, color=DARK)
        # right accent box
        add_rect(slide, Inches(7.9), body_y, Inches(4.9), body_h, deck_color)
        add_text_box(slide, accent_label, Inches(8.0), body_y + Inches(0.15),
                     Inches(4.7), Inches(0.55), font_size=18, bold=True,
                     color=WHITE, align=PP_ALIGN.CENTER)
        add_text_box(slide, accent_text, Inches(8.05), body_y + Inches(0.75),
                     Inches(4.7), body_h - Inches(0.9), font_size=18,
                     color=WHITE, align=PP_ALIGN.LEFT, font_name='Arial')
    else:
        add_bullet_box(slide, bullets, Inches(0.45), body_y, Inches(12.4), body_h,
                       font_size=21, color=DARK)

    # slide counter + footer
    add_footer(slide)
    if slide_num and total:
        add_text_box(slide, f'{slide_num} / {total}',
                     W - Inches(1.1), H - Inches(0.42), Inches(0.9), Inches(0.38),
                     font_size=12, color=WHITE, align=PP_ALIGN.RIGHT, font_name='Arial')
    return slide

def challenge_slide(title, items, message, deck_color=MAROON):
    """Special slide for challenges / parent action."""
    slide = prs.slides.add_slide(blank_layout)
    add_rect(slide, 0, 0, W, H, LIGHT_MAROON)
    add_rect(slide, 0, 0, W, Inches(1.35), deck_color)
    add_rect(slide, 0, 0, Inches(0.22), H, deck_color)
    add_text_box(slide, title, Inches(0.45), Inches(0.12),
                 Inches(11.5), Inches(1.1), font_size=38, bold=True,
                 color=WHITE, align=PP_ALIGN.LEFT)
    add_logo(slide)

    # checklist boxes
    y = Inches(1.6)
    for item in items:
        add_rect(slide, Inches(0.5), y, Inches(0.5), Inches(0.42), WHITE)
        add_text_box(slide, item, Inches(1.15), y, Inches(11.5), Inches(0.55),
                     font_size=20, color=DARK, font_name='Arial')
        y += Inches(0.6)

    # bottom message box
    add_rect(slide, Inches(0.4), H - Inches(1.55), Inches(12.5), Inches(1.1), deck_color)
    add_text_box(slide, message, Inches(0.55), H - Inches(1.52),
                 Inches(12.2), Inches(1.05), font_size=19, bold=True,
                 color=WHITE, align=PP_ALIGN.CENTER)
    add_footer(slide)
    return slide

def comparison_slide(title, subtitle, left_label, left_items, right_label, right_items,
                     deck_color=MAROON, left_color=None, right_color=None):
    slide = prs.slides.add_slide(blank_layout)
    add_rect(slide, 0, 0, W, H, OFF_WHITE)
    add_rect(slide, 0, 0, W, Inches(1.35), deck_color)
    add_rect(slide, 0, 0, Inches(0.22), H, deck_color)
    add_text_box(slide, title, Inches(0.45), Inches(0.12),
                 Inches(11.2), Inches(0.80), font_size=38, bold=True,
                 color=WHITE, align=PP_ALIGN.LEFT)
    if subtitle:
        add_text_box(slide, subtitle, Inches(0.45), Inches(0.88),
                     Inches(11.2), Inches(0.42), font_size=20,
                     color=WHITE, align=PP_ALIGN.LEFT, font_name='Arial')
    add_logo(slide)

    col_y = Inches(1.5)
    col_h = Inches(5.3)
    lc = left_color  or MAROON
    rc = right_color or GREEN

    # Left column
    add_rect(slide, Inches(0.4), col_y, Inches(6.0), col_h, lc)
    add_text_box(slide, left_label, Inches(0.5), col_y + Inches(0.1),
                 Inches(5.8), Inches(0.6), font_size=24, bold=True,
                 color=WHITE, align=PP_ALIGN.CENTER)
    add_bullet_box(slide, left_items, Inches(0.55), col_y + Inches(0.8),
                   Inches(5.7), col_h - Inches(1.0), font_size=19,
                   color=WHITE, font_name='Arial', bullet_char='✗')

    # Right column
    add_rect(slide, Inches(6.9), col_y, Inches(6.0), col_h, rc)
    add_text_box(slide, right_label, Inches(7.0), col_y + Inches(0.1),
                 Inches(5.8), Inches(0.6), font_size=24, bold=True,
                 color=WHITE, align=PP_ALIGN.CENTER)
    add_bullet_box(slide, right_items, Inches(7.05), col_y + Inches(0.8),
                   Inches(5.7), col_h - Inches(1.0), font_size=19,
                   color=WHITE, font_name='Arial', bullet_char='✓')

    add_footer(slide)
    return slide

def parent_slide(deck_color=GREEN):
    slide = prs.slides.add_slide(blank_layout)
    add_rect(slide, 0, 0, W, H, LIGHT_GREEN)
    add_rect(slide, 0, 0, W, Inches(1.35), deck_color)
    add_rect(slide, 0, 0, Inches(0.22), H, MAROON)

    add_text_box(slide, '🏠  TEACH YOUR PARENTS TONIGHT!',
                 Inches(0.45), Inches(0.08), Inches(11.5), Inches(1.1),
                 font_size=36, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_logo(slide)

    # Left: what to say
    add_rect(slide, Inches(0.4), Inches(1.5), Inches(6.0), Inches(3.7), MAROON)
    add_text_box(slide, '💬  Say This Tonight:',
                 Inches(0.5), Inches(1.6), Inches(5.8), Inches(0.55),
                 font_size=21, bold=True, color=WHITE)
    scripts = [
        '"Mom/Dad — did you know one fast food dinner\ncosts $50? I can help us make it at home for $10!"',
        '"Can I get my cut of what we save and put it\nin my piggy bank? I promise to save it!"',
        '"Let\'s count how much we saved this month\ntogether — I\'ll keep track!"',
    ]
    y = Inches(2.25)
    for s in scripts:
        add_rect(slide, Inches(0.5), y, Inches(5.8), Inches(0.8), WHITE)
        add_text_box(slide, s, Inches(0.58), y + Inches(0.04),
                     Inches(5.6), Inches(0.75), font_size=15,
                     color=DARK, font_name='Arial')
        y += Inches(0.95)

    # Right: savings table + piggy bank deal
    add_rect(slide, Inches(6.9), Inches(1.5), Inches(6.0), Inches(3.7), deck_color)
    add_text_box(slide, '🐷  The Piggy Bank Deal:',
                 Inches(7.0), Inches(1.6), Inches(5.8), Inches(0.55),
                 font_size=21, bold=True, color=WHITE)
    rows = [
        ('Meal cooked at home', 'Your Cut'),
        ('Saves ~$35 vs fast food', '$3 – $5 in piggy bank'),
        ('3 meals/week',           'Up to $15/week saved'),
        ('1 month',                'Up to $60 in your jar!'),
        ('1 year',                 'Up to $720 saved! 🎉'),
    ]
    ry = Inches(2.25)
    for i, (left, right) in enumerate(rows):
        bg = RGBColor(0x1e, 0x3d, 0x1c) if i % 2 == 0 else deck_color
        add_rect(slide, Inches(7.0), ry, Inches(5.8), Inches(0.56), bg)
        add_text_box(slide, left, Inches(7.08), ry + Inches(0.05),
                     Inches(3.2), Inches(0.5), font_size=16,
                     color=WHITE, font_name='Arial', bold=(i==0))
        add_text_box(slide, right, Inches(10.3), ry + Inches(0.05),
                     Inches(2.4), Inches(0.5), font_size=16,
                     color=GOLD if i > 0 else WHITE, font_name='Arial', bold=True)
        ry += Inches(0.59)

    # Bottom promise banner
    add_rect(slide, Inches(0.4), Inches(5.38), Inches(12.5), Inches(0.85), MAROON)
    add_text_box(slide,
        '👨‍👩‍👧  Parent Promise: "Every time we cook at home instead of eating out, I will give you your cut for your piggy bank."',
        Inches(0.55), Inches(5.4), Inches(12.2), Inches(0.82),
        font_size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    add_footer(slide, 'Pete Brown Jr. Tennis Program — What we teach on the court, we teach for life.')
    return slide

# ══════════════════════════════════════════════════════════════════════════════
# DECK 1 — WHY WE AVOID FAST FOOD
# ══════════════════════════════════════════════════════════════════════════════
title_slide('THE FAST FOOD TRAP 🍔🚫',
            "It's fast to eat. The damage is slow.",
            "What your favourite meal is actually doing to your body",
            1, MAROON)

content_slide(
    'You\'d Be Shocked What\'s Inside 😱',
    'Fast food is engineered to make you addicted — not healthy.',
    [
        '🧂  Salt: 2,000–3,000 mg — your whole DAY\'s limit in ONE meal',
        '🍬  Sugar: Hidden in buns, sauces & drinks — up to 20 teaspoons!',
        '🛢️  Fat: Mostly unhealthy trans fats that clog your arteries',
        '🧪  Chemicals: Artificial flavours, colours & preservatives',
        '🔥  Calories: One combo meal = a kid\'s ENTIRE daily need',
    ],
    accent_label='⚡ REMEMBER',
    accent_text='Food scientists design fast food to make you want MORE — not to make you HEALTHY.\n\nYou are not weak.\nThe food is designed to hook you.',
    deck_color=MAROON, slide_num=2, total=6
)

content_slide(
    'Your Body After Fast Food 🤢',
    'Here\'s what happens hour by hour...',
    [
        '0–30 min   → Blood sugar spikes — you feel a rush of energy',
        '30–60 min  → Blood sugar CRASHES — tired, foggy, irritable',
        '1–2 hours  → Hungry AGAIN even after 1,500 calories',
        'Weeks later → Arteries begin to harden from excess fat & salt',
        'Years later  → Risk of heart disease & diabetes climbs fast',
    ],
    accent_label='📊 STAT',
    accent_text='Kids who eat fast food 3+ times a week are TWICE as likely to be overweight by age 12.',
    deck_color=MAROON, slide_num=3, total=6
)

content_slide(
    'Fast Food Hurts Your Grades Too 🧠📉',
    'What you eat directly affects how well you think.',
    [
        '🔴  High sugar + fat meals reduce blood flow to the brain',
        '🔴  Kids who eat fast food regularly score lower on tests',
        '🔴  The crash after eating makes it hard to focus in class',
        '🔴  Fast food has almost none of the nutrients brains need to grow',
        '🟢  After a home-cooked meal → alert, focused, energised',
        '🟢  Omega-3s (fish, nuts) = brain food. Fast food has almost NONE.',
    ],
    deck_color=MAROON, slide_num=4, total=6
)

content_slide(
    'Fast Food Is Designed to Hook You 🪝',
    'This isn\'t about willpower. The food is engineered to be addictive.',
    [
        '💡  Food companies find the "bliss point" — the perfect mix of salt,',
        '     sugar & fat that makes your brain release dopamine',
        '⚠️  Signs you might be hooked:',
        '     → You crave it even when not hungry',
        '     → You feel moody without it',
        '     → You feel guilty but do it again anyway',
        '✅  Good news: Within 2 WEEKS of cutting fast food,',
        '     cravings drop and healthy food tastes better again!',
    ],
    deck_color=MAROON, slide_num=5, total=6
)

comparison_slide(
    'Better Choices, Better You ✅',
    'You don\'t have to be perfect. You have to be CONSISTENT.',
    '❌  Fast Food',
    [
        '10+ teaspoons of sugar',
        'Over a day\'s worth of salt',
        'Energy crash in 30 min',
        'Hungry again in 1 hour',
        'Hurts focus at school',
        'Expensive long-term',
    ],
    '✅  Smart Swaps',
    [
        'Banana + peanut butter',
        'Hard boiled egg + crackers',
        'Cheese + apple slices',
        'Trail mix (nuts & dried fruit)',
        'Leftovers from home',
        'Saves your family $$$',
    ],
    deck_color=MAROON
)

# ══════════════════════════════════════════════════════════════════════════════
# DECK 2 — WHY WE AVOID TOO MUCH SUGAR
# ══════════════════════════════════════════════════════════════════════════════
title_slide('THE SUGAR SECRET 🍬😬',
            'It tastes amazing. Here\'s what it\'s actually doing.',
            'The truth that candy companies don\'t want you to know',
            2, GREEN)

content_slide(
    'See the Sugar 👀',
    'Daily limit for kids: 6 teaspoons (25g). Look what\'s hiding in your drinks!',
    [
        '🥤  1 can of soda          →  10 teaspoons of sugar',
        '📦  1 juice box            →  6 teaspoons  (almost as bad!)',
        '🍶  Flavoured yogurt       →  5 teaspoons',
        '⚡  1 sports drink          →  8 teaspoons',
        '🍫  1 chocolate milk       →  4 teaspoons',
        '🍪  1 granola bar          →  4 teaspoons',
        '',
        '😱  Average American kid eats 19 teaspoons/day — 3× the limit!',
    ],
    accent_label='💧 BEST DRINK',
    accent_text='#1  Water\n(zero sugar, zero cost)\n\n#2  Plain milk\n(calcium + protein)\n\n#3  Unsweetened tea\n(zero sugar)',
    deck_color=GREEN, slide_num=1, total=6
)

content_slide(
    'The Sugar Rollercoaster 🎢',
    'Every time you eat a lot of sugar your body goes on a wild ride.',
    [
        '⬆️  THE SPIKE:',
        '   Sugar hits blood → glucose shoots up fast',
        '   You feel: hyper, excited, can\'t sit still',
        '',
        '⬇️  THE CRASH:',
        '   Insulin overreacts → blood sugar drops too low',
        '   You feel: tired, grumpy, headachey, can\'t focus',
        '',
        '🔁  THE CRAVING:',
        '   You want MORE sugar to feel normal again',
        '   This is exactly how sugar keeps you hooked.',
    ],
    deck_color=GREEN, slide_num=2, total=6
)

content_slide(
    'What Too Much Sugar Does Over Time ⏳',
    'A little occasionally is fine. Daily excess causes real damage.',
    [
        '🦷  TEETH: Sugar feeds bacteria → acid → cavities',
        '💪  BODY: Excess sugar is stored as fat (especially belly fat)',
        '💉  DIABETES: Type 2 diabetes now showing up in kids as young as 10',
        '❤️  HEART: Causes inflammation in blood vessels',
        '🛡️  IMMUNITY: Weakens your immune system for up to 5 hours after a sugar binge',
        '🧠  BRAIN: Linked to anxiety, poor memory & difficulty learning',
        '',
        '📈  1900s: Americans ate 4 lbs sugar/year',
        '     Today: 150+ lbs sugar/year per person',
    ],
    deck_color=GREEN, slide_num=3, total=6
)

content_slide(
    'Sugar in Disguise 🕵️',
    'Food companies hide sugar under 56 different names on labels.',
    [
        '🔍  Watch for these names on ingredient lists:',
        '   → High fructose corn syrup',
        '   → Dextrose, maltose, sucrose, fructose',
        '   → Evaporated cane juice / cane sugar',
        '   → Agave nectar / brown rice syrup',
        '   → Fruit juice concentrate',
        '',
        '📏  RULE: If ANY form of sugar is in the first 3 ingredients — PUT IT BACK.',
        '',
        '✅  Safe test: Sugar grams ÷ 4 = teaspoons. Under 6g per serving = OK.',
    ],
    deck_color=GREEN, slide_num=4, total=6
)

content_slide(
    'Beat Sugar Cravings Like a Pro 🏆',
    'You don\'t have to go cold turkey. Just crowd it out — week by week.',
    [
        '📅  Week 1: Swap ONE sugary drink a day for water',
        '📅  Week 2: Replace ONE sugary snack with fresh fruit',
        '📅  Week 3: Read ONE food label before buying — check sugar grams',
        '📅  Week 4: Try ONE full day with zero added sugar',
        '',
        '⚡  When cravings hit RIGHT NOW:',
        '   → Drink a full glass of water — wait 10 minutes',
        '   → Eat a small handful of nuts — kills cravings fast',
        '   → Brush your teeth — mint flavour stops cravings',
        '   → Go outside and move — exercise beats sugar cravings in minutes',
    ],
    accent_label='🌟 RESULT',
    accent_text='After just 2 weeks of less sugar:\n\n✓  More steady energy\n✓  Better mood\n✓  Food tastes MORE flavourful\n✓  Easier to focus',
    deck_color=GREEN, slide_num=5, total=6
)

challenge_slide(
    '🎯  Your 5-Day Sugar Challenge',
    [
        'Day 1 — Replace soda/juice with water all day',
        'Day 2 — Read the label on 3 things in your fridge',
        'Day 3 — Count teaspoons of sugar in everything you drink',
        'Day 4 — Choose a fruit instead of a sweet snack',
        'Day 5 — Tell your parent ONE thing you learned about sugar',
    ],
    'After 5 days write down: How do you feel compared to Day 1? More energy? Better mood? 🌟',
    deck_color=GREEN
)

# ══════════════════════════════════════════════════════════════════════════════
# DECK 3 — FAST FOOD vs HOME COOKING + SAVE MONEY
# ══════════════════════════════════════════════════════════════════════════════
title_slide('FAST FOOD vs HOME COOKING 🍔⚔️🍳',
            'Which one wins for your health AND your wallet?',
            'The answer — and the savings — might surprise you',
            3, GOLD)

content_slide(
    'Fast Food Is Expensive — Do The Maths 💸',
    'It feels cheap per visit. The yearly total will shock you.',
    [
        '🧒  ONE PERSON eating fast food:',
        '   1 meal = $10–$14    |   3×/week = $35–$42',
        '   1 month = $140–$168  |   1 YEAR = $1,680–$2,016',
        '',
        '👨‍👩‍👧‍👦  FAMILY OF 4 eating fast food:',
        '   1 meal out = $40–$55  |  3×/week = $120–$165',
        '   1 YEAR = $6,240–$8,580',
        '',
        '💡  That is a family vacation. A used car. A year of college savings.',
        '    And that\'s BEFORE the medical bills from eating unhealthy.',
    ],
    deck_color=GOLD, slide_num=1, total=6
)

content_slide(
    'What Home Cooking Actually Costs 🏡',
    'Cooking at home is cheaper than you think — especially per serving.',
    [
        '🍝  Pasta + sauce + turkey for 4:    $8 total  →  $2 per person',
        '🍗  Chicken + veg + rice for 4:       $12 total →  $3 per person',
        '🥚  Eggs + toast + fruit for 4:        $5 total  →  $1.25 per person',
        '🌮  Bean tacos with toppings for 4:   $10 total →  $2.50 per person',
        '',
        '                        Fast Food     Home Cooked   YOU SAVE',
        '   Breakfast (4 ppl):   $28            $6               $22',
        '   Dinner (4 ppl):       $50            $12              $38',
        '   Per week (×5 dinners): $250          $60              $190',
        '',
        '💰  Annual saving for a family switching to home cooking: $4,000–$6,000',
    ],
    deck_color=GOLD, slide_num=2, total=6
)

content_slide(
    'You Pay Now or You Pay Later 🏥',
    'Cheap unhealthy food now = expensive medical bills later.',
    [
        '🍔  FAST FOOD DIET — long-term costs:',
        '   → Obesity management:        $1,500+/year extra',
        '   → Type 2 diabetes:           $9,000–$13,000/year',
        '   → Heart disease treatment:   $20,000–$50,000+ lifetime',
        '   → Dental work from sugar:    $500–$5,000',
        '',
        '🍳  HOME-COOKED DIET — long-term savings:',
        '   → Fewer doctor visits',
        '   → Lower medication costs',
        '   → More energy at school & sport',
        '   → Stronger immune system = less sick time',
        '',
        '⚖️  Cheap food now + expensive health problems later, OR',
        '    A little more effort now + a longer, healthier, wealthier life.',
    ],
    deck_color=GOLD, slide_num=3, total=6
)

content_slide(
    'Cooking Is a Superpower Anyone Can Learn 👨‍🍳',
    '5 things every kid should know how to make:',
    [
        '🍳  1.  Scrambled eggs — protein in 5 minutes',
        '🍚  2.  Rice in a pot — base for any meal',
        '🥗  3.  A simple salad — chop + dressing = done',
        '🍝  4.  Pasta with sauce — feeds a family for $8',
        '🥪  5.  A proper sandwich — protein + veg + whole grain',
        '',
        '✅  Why cooking matters beyond the money:',
        '   → You control EXACTLY what goes in your food',
        '   → No hidden salt, sugar, or chemicals',
        '   → It\'s a life skill that will save you THOUSANDS as an adult',
        '   → Kids who cook by age 12 eat healthier as adults — proven!',
    ],
    deck_color=GOLD, slide_num=4, total=6
)

comparison_slide(
    '30% Savings Rule 🐷  — Every Time You Get Money',
    'Pay YOURSELF first. Always.',
    '❌  Spend It All',
    [
        'Money is gone immediately',
        'Nothing saved for goals',
        'Emergency = panic',
        'Always need more money',
        'Stressed about spending',
        '$10 earned → $0 left',
    ],
    '✅  The 30% Rule',
    [
        '30% straight into savings jar',
        '10% into giving jar',
        '60% to spend wisely',
        'Emergency fund grows slowly',
        'Watching savings grow feels GREAT',
        '$10 earned → $3 saved every time',
    ],
    deck_color=GOLD,
    left_color=MAROON,
    right_color=GREEN
)

challenge_slide(
    '🏆  Cook It · Track It · Save It This Week!',
    [
        'Step 1 — Pick a simple recipe  (eggs, pasta, rice bowl, or sandwich)',
        'Step 2 — Go to the store with a parent and buy only what you need',
        'Step 3 — Cook it together and write down what it cost',
        'Step 4 — Compare the cost to the same meal at a restaurant',
        'Step 5 — Put 30% of the savings straight into your piggy bank 🐷',
    ],
    'Example: Home pasta dinner = $8.  Same meal at restaurant = $45.  You saved $37 — $11 goes in your piggy bank!',
    deck_color=GOLD
)

# Parent slide
parent_slide(deck_color=GREEN)

prs.save(OUT)
print(f'✅  Saved: {OUT}')
print(f'   Slides: {len(prs.slides)}')
