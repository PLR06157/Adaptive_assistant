# Email-Compatible HTML Guidelines for AI/LLM

When generating HTML for email templates, follow these rules to ensure compatibility across all email clients (Outlook, Gmail, Apple Mail, mobile apps).

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
| Borders | `border`, `border-left`, `border-top`, `border-bottom`, `border-right`, `border-radius` |
| Display | `display: inline-block`, `display: block` |

### 3. Use `<div>` and `<p>` for Layout

Avoid flexbox, grid, or complex positioning. Use simple nested `<div>` elements.

### 4. Specify Font Stacks

Always include fallback fonts:

```html
<!-- Body text -->
<p style="font-family: 'Open Sans', Roboto, Arial, sans-serif;">Text</p>

<!-- Headings -->
<h1 style="font-family: Copperplate, Georgia, 'Times New Roman', serif;">Heading</h1>
```

### 5. Use Full Hex Colors

```html
<!-- CORRECT -->
<p style="color: #2C2C2C;">Text</p>

<!-- WRONG -->
<p style="color: #2C2;">Text</p>
```

### 6. Set Max-Width on Body

Standard email width is 600px:

```html
<body style="max-width: 600px; margin: 0 auto; padding: 20px;">
```

### 7. Style Buttons as Links

```html
<a href="https://example.com" style="display: inline-block; background-color: #B8975A; color: white; padding: 12px 30px; text-decoration: none; border-radius: 4px; font-weight: bold;">
    Button Text
</a>
```

---

## MUST NOT DO

| Rule | Reason |
|------|--------|
| Never use `<style>` blocks | Email clients strip them |
| Never use CSS classes | Classes require `<style>` block |
| Never use `position`, `float`, `flexbox`, or `grid` | Poor/no support |
| Never use `@media` queries | No responsive design via CSS in emails |
| Never use `:hover` pseudo-classes | Not supported inline |
| Never use external stylesheets | Will not load |
| Never use web fonts via `@font-face` | Will not load |
| Never use JavaScript | Blocked by all email clients |

---

## Template Structure

```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Email Subject</title>
</head>
<body style="font-family: 'Open Sans', Arial, sans-serif; line-height: 1.6; color: #2C2C2C; max-width: 600px; margin: 0 auto; padding: 20px; background-color: #F8F6F3;">
    <div style="background-color: #FFFFFF; border-radius: 8px; padding: 30px;">

        <!-- HEADER -->
        <div style="border-bottom: 3px solid #B8975A; padding-bottom: 20px; margin-bottom: 20px;">
            <h1 style="font-family: Copperplate, Georgia, 'Times New Roman', serif; color: #B8975A; margin: 0; font-size: 24px; font-weight: 700;">
                Email Title
            </h1>
            <span style="display: inline-block; background-color: #4CAF50; color: white; padding: 5px 15px; border-radius: 4px; font-weight: bold; font-size: 12px; text-transform: uppercase; margin-top: 10px;">
                Status Badge
            </span>
        </div>

        <!-- GREETING -->
        <p style="font-weight: 400;">Dear Recipient Name,</p>

        <!-- ICON (optional) -->
        <div style="font-size: 48px; text-align: center; margin: 20px 0;">&#9989;</div>

        <!-- MAIN MESSAGE -->
        <p style="font-weight: 400;">Your main message content goes here.</p>

        <!-- INFO SECTION -->
        <div style="margin: 20px 0;">
            <div style="margin: 10px 0; padding: 10px; background-color: #F8F6F3; border-left: 4px solid #B8975A;">
                <strong style="color: #2C2C2C; display: inline-block; min-width: 150px; font-weight: 600;">Label:</strong> Value
            </div>
            <div style="margin: 10px 0; padding: 10px; background-color: #F8F6F3; border-left: 4px solid #B8975A;">
                <strong style="color: #2C2C2C; display: inline-block; min-width: 150px; font-weight: 600;">Another Label:</strong> Another Value
            </div>
        </div>

        <!-- ALERT/NOTE BOX -->
        <div style="background-color: #fff3e0; border-left: 4px solid #FF9800; padding: 15px; margin: 20px 0; border-radius: 4px;">
            <strong style="color: #E65100;">Note:</strong> Important information or call to action.
        </div>

        <!-- SUCCESS BOX (alternative) -->
        <div style="background-color: #e8f5e9; border-left: 4px solid #4CAF50; padding: 15px; margin: 20px 0; border-radius: 4px;">
            <strong style="color: #2e7d32;">Success:</strong> Positive confirmation message.
        </div>

        <!-- ERROR BOX (alternative) -->
        <div style="background-color: #ffebee; border-left: 4px solid #F44336; padding: 15px; margin: 20px 0; border-radius: 4px;">
            <strong style="color: #c62828;">Error:</strong> Error or rejection message.
        </div>

        <!-- CTA BUTTON -->
        <div style="text-align: center;">
            <a href="https://example.com" style="display: inline-block; background-color: #B8975A; color: white; padding: 12px 30px; text-decoration: none; border-radius: 4px; margin: 20px 0; font-weight: bold;">
                Call to Action
            </a>
        </div>

        <!-- CLOSING -->
        <p style="font-weight: 400;">
            Closing message here.<br>
            <strong>Adaptive</strong>
        </p>

        <!-- FOOTER -->
        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #C5A572; font-size: 12px; color: #9E9E9E; text-align: center;">
            <p>This is an automated notification.</p>
            <p>Please do not reply to this email.</p>
        </div>
    </div>
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
| Card border-radius | `8px` | Container elements, cards |
| Button border-radius | `4px` | CTAs, badges, small elements |
| Subtle shadow | `0px 2px 6px rgba(0, 0, 0, 0.10)` | Cards, elevated elements |

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

## Testing Recommendations

1. **Test in multiple clients**: Outlook (Windows), Gmail (Web), Apple Mail, mobile apps
2. **Check dark mode**: Some clients invert colors
3. **Verify links**: Ensure all `href` attributes are absolute URLs
4. **Check images**: Use absolute URLs, add `alt` text, set explicit `width`/`height`
