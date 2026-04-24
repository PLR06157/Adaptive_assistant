# Email-Compatible HTML Guidelines for AI/LLM

When generating HTML for email templates, follow these rules to ensure compatibility across all email clients (Outlook, Gmail, Apple Mail, mobile apps).

---

## CRITICAL: Table-Based Layout

**All email layout MUST use `<table>` elements.** Outlook (Windows) uses Microsoft Word's rendering engine, which does not support `<div>`-based layout, `max-width` on body, flexbox, grid, or CSS-only centering. Using `<div>` for layout will break rendering in Outlook.

### Required Structure

Every email must use this wrapper pattern:

```html
<body style="margin: 0; padding: 0; background-color: #F8F6F3; font-family: 'Open Sans', Arial, sans-serif; line-height: 1.6; color: #2C2C2C;">
    <!-- Outer wrapper table - handles background color and centering -->
    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8F6F3;">
        <tr>
            <td align="center" style="padding: 20px 0;">
                <!-- Inner content table - fixed 600px width -->
                <table role="presentation" width="600" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border-radius: 8px;">

                    <!-- Each content section is a <tr> -->
                    <tr>
                        <td style="padding: 20px 24px;">
                            Content here
                        </td>
                    </tr>

                </table>
            </td>
        </tr>
    </table>
</body>
```

### Table Rules

| Rule | Details |
|------|---------|
| Always use `role="presentation"` | Prevents screen readers from treating layout tables as data tables |
| Always set `border="0" cellpadding="0" cellspacing="0"` | Resets default table spacing across all clients |
| Use `width="600"` attribute on inner table | The HTML attribute works more reliably than CSS `width` in Outlook |
| Each content block = one `<tr><td>...</td></tr>` | Sections (header, body, CTA, footer) should be separate rows |
| Use `align="center"` on wrapper `<td>` | Centers the inner table; `margin: 0 auto` does NOT work in Outlook |
| Nest tables for sub-layouts | For side-by-side elements, buttons, badges - use inner tables, not divs |

---

## MUST DO

### 1. Use Inline CSS Only

Place all styles directly in `style=""` attributes on elements. Never use `<style>` blocks in `<head>`.

```html
<!-- CORRECT -->
<p style="color: #333; font-size: 16px;">Text</p>

<!-- WRONG -->
<style>
  .text { color: #333; font-size: 16px; }
</style>
<p class="text">Text</p>
```

### 2. Keep CSS Simple

Use only basic CSS properties:

| Category | Supported Properties |
|----------|---------------------|
| Typography | `font-family`, `font-size`, `font-weight`, `color`, `line-height`, `text-align`, `text-decoration`, `text-transform` |
| Box Model | `background-color`, `padding`, `margin`, `width`, `max-width`, `min-width` |
| Borders | `border`, `border-left`, `border-top`, `border-bottom`, `border-right` |
| Display | `display: inline-block`, `display: block` |

**Outlook limitations on CSS:**

| Property | Outlook Support |
|----------|----------------|
| `border-radius` | Ignored (renders as square corners) |
| `max-width` | Ignored on `<body>` and `<div>` (use `width` attribute on `<table>` instead) |
| `background-color` on `<a>` | Ignored (wrap in `<table><td>` with background instead) |
| `margin` | Partial support; use `padding` on `<td>` when possible |
| `line-height` | Use px values, not unitless numbers |

### 3. Specify Font Stacks

Always include fallback fonts. Repeat `font-family` on each `<td>` that contains text, as font inheritance is unreliable across email clients.

```html
<!-- Body text -->
<td style="font-family: 'Open Sans', Roboto, Arial, sans-serif;">
    <p style="font-size: 14px;">Text</p>
</td>

<!-- Headings -->
<h1 style="font-family: Copperplate, Georgia, 'Times New Roman', serif;">Heading</h1>
```

### 4. Use Full Hex Colors

```html
<!-- CORRECT -->
<p style="color: #2C2C2C;">Text</p>

<!-- WRONG -->
<p style="color: #2C2;">Text</p>
```

### 5. Style Buttons Using Table Cells

Wrapping buttons in a `<table>` ensures the background color renders in Outlook. Never rely on `background-color` on `<a>` alone.

```html
<!-- CORRECT - works in Outlook -->
<table role="presentation" border="0" cellpadding="0" cellspacing="0">
    <tr>
        <td align="center" style="background-color: #B8975A; border-radius: 4px;">
            <a href="https://example.com" style="display: inline-block; color: #FFFFFF; padding: 12px 30px; text-decoration: none; font-weight: bold; font-size: 14px; font-family: 'Open Sans', Arial, sans-serif;">
                Button Text
            </a>
        </td>
    </tr>
</table>

<!-- WRONG - background disappears in Outlook -->
<a href="https://example.com" style="display: inline-block; background-color: #B8975A; color: white; padding: 12px 30px; text-decoration: none;">
    Button Text
</a>
```

### 6. Images

| Rule | Details |
|------|---------|
| Always use absolute URLs | `src="https://..."` - relative paths break when sent as email |
| Set explicit `width` and `height` attributes | Prevents layout shift and ensures sizing in Outlook |
| Add `style="display: block; border: 0;"` | Removes unwanted gaps and blue borders on linked images |
| Always include `alt` text | Displayed when images are blocked (many corporate clients block by default) |
| Use `line-height: 0` on parent `<td>` for full-width images | Prevents whitespace gaps below images |

```html
<tr>
    <td style="padding: 0; line-height: 0;">
        <img src="https://example.com/header.jpg" alt="Header" width="600" height="200" style="display: block; width: 100%; height: auto; border: 0;">
    </td>
</tr>
```

### 7. Photo Galleries (Multi-Column Image Grids)

Outlook on Windows (Word rendering engine) ignores `width="100%"` and `height:auto` on `<img>` tags
inside percentage-based `<td>` cells. It renders images at their natural pixel size, then clips them.
The clip typically cuts off **heads** in people photos.

**Always run `prepare_email.py` before sending any template that contains local gallery images.**
The script reads actual pixel dimensions with Pillow and stamps explicit `width`/`height` attributes
on every `<img>`, and adds `valign="top"` so any residual clipping hits the bottom, not the top.

```bash
python3 mailing/prepare_email.py --template mailing/sets/<folder>/template.html
```

**Rules when writing gallery HTML:**

| Rule | Details |
|------|---------|
| Use `width="N%"` or `width="Npx"` on `<td>` | The script needs this to calculate display width |
| Use `padding:Npx` on `<td>` for gaps | The script subtracts horizontal padding automatically |
| Do NOT set explicit `width`/`height` on gallery `<img>` | Leave it to the script; it recalculates each run |
| Icons/logos with fixed sizes: set `width="N" height="N"` explicitly | The script skips images that already have both attributes |
| Use `valign="top"` on gallery `<td>` | The script adds this; include it by hand if writing without the script |

**Gallery template snippet (write it like this; the script fills in dimensions):**

```html
<!-- 4-column photo grid -->
<tr>
    <td style="padding:0; line-height:0;">
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
                <td width="25%" style="padding:1px; line-height:0;">
                    <img src="1.png" alt="" width="100%" style="display:block; border:0; width:100%; height:auto;">
                </td>
                <td width="25%" style="padding:1px; line-height:0;">
                    <img src="2.png" alt="" width="100%" style="display:block; border:0; width:100%; height:auto;">
                </td>
                <!-- … -->
            </tr>
        </table>
    </td>
</tr>
```

After running `prepare_email.py`, each `<img>` will have exact pixel dimensions and each `<td>` will have `valign="top"`.

### 8. Use HTML Entities for Special Characters

Avoid raw Unicode where possible. Use HTML entities for reliability:

```html
<!-- CORRECT -->
Krak&oacute;w
&#128578; <!-- smiley -->

<!-- RISKY -->
Kraków  <!-- may break in some encodings -->
```

---

## MUST NOT DO

| Rule | Reason |
|------|--------|
| Never use `<style>` blocks | Email clients strip them |
| Never use CSS classes | Classes require `<style>` block |
| Never use `<div>` for layout structure | Outlook ignores div-based layout; use `<table>` rows |
| Never use `position`, `float`, `flexbox`, or `grid` | Poor/no support |
| Never use `@media` queries | No responsive design via CSS in emails |
| Never use `:hover` pseudo-classes | Not supported inline |
| Never use external stylesheets | Will not load |
| Never use web fonts via `@font-face` | Will not load |
| Never use JavaScript | Blocked by all email clients |
| Never use `max-width` on body for centering | Use wrapper `<table>` with `width` attribute instead |
| Never use `background-color` on `<a>` for buttons | Outlook strips it; wrap in `<table><td>` |
| Never use relative image paths | Images must be hosted with absolute URLs |

**Note on `<div>` and `<p>`:** You may still use `<div>` and `<p>` *inside* table cells for text-level formatting (e.g., paragraphs, inline grouping). Just never rely on `<div>` for structural layout (width, centering, sections).

---

## Full Template Structure

```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Email Subject</title>
</head>
<body style="margin: 0; padding: 0; background-color: #F8F6F3; font-family: 'Open Sans', Arial, sans-serif; line-height: 1.6; color: #2C2C2C;">
    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8F6F3;">
        <tr>
            <td align="center" style="padding: 20px 0;">
                <table role="presentation" width="600" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border-radius: 8px;">

                    <!-- HEADER IMAGE (optional) -->
                    <tr>
                        <td style="padding: 0; line-height: 0;">
                            <img src="https://example.com/header.jpg" alt="Header" width="600" height="200" style="display: block; width: 100%; height: auto; border: 0; border-radius: 8px 8px 0 0;">
                        </td>
                    </tr>

                    <!-- TITLE -->
                    <tr>
                        <td style="padding: 20px 24px 0 24px;">
                            <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td style="border-bottom: 3px solid #B8975A; padding-bottom: 20px; text-align: center;">
                                        <h1 style="font-family: Copperplate, Georgia, 'Times New Roman', serif; color: #B8975A; margin: 0; font-size: 24px; font-weight: 700;">
                                            Email Title
                                        </h1>
                                        <!-- Optional badge -->
                                        <table role="presentation" border="0" cellpadding="0" cellspacing="0" align="center" style="margin-top: 10px;">
                                            <tr>
                                                <td style="background-color: #4CAF50; color: #FFFFFF; padding: 5px 14px; border-radius: 4px; font-weight: bold; font-size: 12px; text-transform: uppercase; font-family: 'Open Sans', Arial, sans-serif;">
                                                    Status Badge
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>

                    <!-- GREETING -->
                    <tr>
                        <td style="padding: 20px 24px 0 24px; font-family: 'Open Sans', Arial, sans-serif;">
                            <p style="font-weight: 400; font-size: 14px; margin: 0 0 12px 0;">Dear Recipient Name,</p>
                        </td>
                    </tr>

                    <!-- ICON (optional) -->
                    <tr>
                        <td align="center" style="padding: 10px 24px; font-size: 48px;">
                            &#9989;
                        </td>
                    </tr>

                    <!-- MAIN MESSAGE -->
                    <tr>
                        <td style="padding: 0 24px; font-family: 'Open Sans', Arial, sans-serif;">
                            <p style="font-weight: 400; font-size: 14px; margin: 0 0 12px 0;">Your main message content goes here.</p>
                        </td>
                    </tr>

                    <!-- INFO SECTION -->
                    <tr>
                        <td style="padding: 10px 24px;">
                            <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td style="padding: 10px; background-color: #F8F6F3; border-left: 4px solid #B8975A; font-family: 'Open Sans', Arial, sans-serif; font-size: 14px;">
                                        <strong style="color: #2C2C2C; font-weight: 600;">Label:</strong> Value
                                    </td>
                                </tr>
                                <tr><td style="height: 8px;"></td></tr>
                                <tr>
                                    <td style="padding: 10px; background-color: #F8F6F3; border-left: 4px solid #B8975A; font-family: 'Open Sans', Arial, sans-serif; font-size: 14px;">
                                        <strong style="color: #2C2C2C; font-weight: 600;">Another Label:</strong> Another Value
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>

                    <!-- ALERT/NOTE BOX -->
                    <tr>
                        <td style="padding: 10px 24px;">
                            <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td style="background-color: #fff3e0; border-left: 4px solid #FF9800; padding: 15px; border-radius: 4px; font-family: 'Open Sans', Arial, sans-serif; font-size: 14px;">
                                        <strong style="color: #E65100;">Note:</strong> Important information or call to action.
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>

                    <!-- SUCCESS BOX (alternative) -->
                    <tr>
                        <td style="padding: 10px 24px;">
                            <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td style="background-color: #e8f5e9; border-left: 4px solid #4CAF50; padding: 15px; border-radius: 4px; font-family: 'Open Sans', Arial, sans-serif; font-size: 14px;">
                                        <strong style="color: #2e7d32;">Success:</strong> Positive confirmation message.
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>

                    <!-- ERROR BOX (alternative) -->
                    <tr>
                        <td style="padding: 10px 24px;">
                            <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td style="background-color: #ffebee; border-left: 4px solid #F44336; padding: 15px; border-radius: 4px; font-family: 'Open Sans', Arial, sans-serif; font-size: 14px;">
                                        <strong style="color: #c62828;">Error:</strong> Error or rejection message.
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>

                    <!-- CTA BUTTON -->
                    <tr>
                        <td align="center" style="padding: 10px 24px 20px 24px;">
                            <table role="presentation" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td align="center" style="background-color: #B8975A; border-radius: 4px;">
                                        <a href="https://example.com" style="display: inline-block; color: #FFFFFF; padding: 12px 30px; text-decoration: none; font-weight: bold; font-size: 14px; font-family: 'Open Sans', Arial, sans-serif;">
                                            Call to Action
                                        </a>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>

                    <!-- CLOSING -->
                    <tr>
                        <td style="padding: 0 24px 20px 24px; font-family: 'Open Sans', Arial, sans-serif;">
                            <p style="font-weight: 400; font-size: 14px; margin: 0;">
                                Closing message here.<br>
                                <strong>Adaptive</strong>
                            </p>
                        </td>
                    </tr>

                    <!-- FOOTER -->
                    <tr>
                        <td style="padding: 16px 24px 22px 24px; text-align: center; font-size: 12px; color: #9E9E9E; border-top: 1px solid #C5A572; font-family: 'Open Sans', Arial, sans-serif;">
                            <p style="margin: 0;">This is an automated notification.</p>
                            <p style="margin: 8px 0 0 0;">Please do not reply to this email.</p>
                        </td>
                    </tr>

                </table>
            </td>
        </tr>
    </table>
</body>
</html>
```

---

## Brand Colors (Adaptive)

### Core Brand Colors

| Name | Hex Code | Usage |
|------|----------|-------|
| Primary Gold | `#B8975A` | CTAs, headers, accent borders, highlights |
| Light Gold | `#C5A572` | Highlights, secondary accents, footer borders |
| Dark Gold | `#9A7D42` | Depth, shadows, hover states |
| Professional Black | `#000000` | Secondary brand color, strong emphasis |
| Charcoal | `#2C2C2C` | Body text, subtle elements |
| Pure White | `#FFFFFF` | Backgrounds, contrast areas |
| Light Beige | `#F8F6F3` | Soft backgrounds, alternating sections |

### Accent & Status Colors

| Name | Hex Code | Usage |
|------|----------|-------|
| Success Green | `#4CAF50` | Approved, success states |
| Error Red | `#F44336` | Rejected, error states |
| Warning Orange | `#FF9800` | Pending, warnings |
| Info Blue | `#2196F3` | Informational, neutral |
| Neutral Gray | `#9E9E9E` | Draft, disabled, footer text |

---

## Brand Typography

### Font Stacks

| Purpose | Font Stack |
|---------|-----------|
| Headings | `Copperplate, Georgia, 'Times New Roman', serif` |
| Body Text | `'Open Sans', Roboto, Arial, sans-serif` |

### Text Hierarchy

| Element | Weight | Characteristics |
|---------|--------|-----------------|
| Headings (h1, h2) | `700` (Bold) | Professional, authoritative |
| Subheadings | `600` (Semi-bold) | Clear hierarchy |
| Body Text | `400` (Regular) | Clean, readable |
| Labels/Strong | `600` | Emphasis within body |

---

## Brand Visual Style

### Design Values

| Property | Value | Usage |
|----------|-------|-------|
| Card border-radius | `8px` | Container elements (decorative only, ignored in Outlook) |
| Button border-radius | `4px` | CTAs, badges (decorative only, ignored in Outlook) |

### Design Principles

- **Minimalism**: Elegant, clean layouts with generous white space
- **Gold accents**: Use sparingly as premium accent, not overwhelming
- **High contrast**: Ensure readability on mobile devices
- **Professional tone**: B2B aesthetic, corporate sophistication

---

## Color Reference

### Status Colors

| Status | Background | Border | Text |
|--------|------------|--------|------|
| Success/Approved | `#e8f5e9` | `#4CAF50` | `#2e7d32` |
| Warning/Pending | `#fff3e0` | `#FF9800` | `#E65100` |
| Error/Rejected | `#ffebee` | `#F44336` | `#c62828` |
| Info/Neutral | `#e3f2fd` | `#2196F3` | `#1565c0` |
| Special/Admin | `#f3e5f5` | `#9C27B0` | `#7b1fa2` |

### Badge Colors

| Type | Background Color |
|------|-----------------|
| Approved/Success | `#4CAF50` |
| Pending/Warning | `#FF9800` |
| Rejected/Error | `#F44336` |
| Info/Submitted | `#2196F3` |
| Closed/Complete | `#2196F3` |
| Withdrawn | `#FF9800` |
| Admin Action | `#9C27B0` |
| Draft | `#9E9E9E` |

---

## Common HTML Entities for Icons

| Icon | Entity | Description |
|------|--------|-------------|
| &#9989; | `&#9989;` | Checkmark |
| &#10060; | `&#10060;` | Cross/X |
| &#128197; | `&#128197;` | Calendar |
| &#9200; | `&#9200;` | Clock |
| &#128188; | `&#128188;` | Briefcase |
| &#128221; | `&#128221;` | Memo/Document |
| &#128465; | `&#128465;` | Trash |
| &#8630; | `&#8630;` | Undo arrow |
| &#128176; | `&#128176;` | Money bag |
| &#128181; | `&#128181;` | Money with wings |
| &#128196; | `&#128196;` | Document |
| &#128274; | `&#128274;` | Lock |

---

## Deliverability Tips

These won't fix server-level blocking but help avoid spam filters:

| Tip | Details |
|-----|---------|
| Avoid ALL CAPS in body text | "FREE OF CHARGE" triggers spam scoring; use "Free of charge" |
| Maintain good text-to-image ratio | Emails that are mostly images get flagged; keep meaningful text |
| Ensure clean, valid HTML | Broken tags, unclosed elements raise spam scores |
| Use absolute URLs for all links | Relative links are suspicious to filters |
| Include unsubscribe link | Required by CAN-SPAM / GDPR; missing it flags the email |
| Don't use URL shorteners | Shortened links are heavily penalized by spam filters |
| Avoid spam trigger phrases | "Act now", "Limited time", "Click here", excessive exclamation marks |

**Note:** Server-level blocking (SPF/DKIM/DMARC failures, IP reputation) is separate from HTML quality. Ensure your sending domain's DNS records authorize your email platform (e.g., GetResponse).

---

## Testing Recommendations

1. **Test in multiple clients**: Outlook (Windows), Gmail (Web), Apple Mail, mobile apps
2. **Check dark mode**: Some clients invert colors
3. **Verify links**: Ensure all `href` attributes are absolute URLs
4. **Check images**: Use absolute URLs, add `alt` text, set explicit `width`/`height`
5. **Validate HTML**: Ensure all tags are properly closed and nested
6. **Send test emails**: Use tools like Litmus or Email on Acid to preview across 90+ clients
