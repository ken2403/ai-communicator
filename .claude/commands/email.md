# Email Reply Generator

Convert any rough/casual input into a polished, formal business email.

## Language Options
- `-e`: Output in English
- `-j`: Output in Japanese (default)

## Input
$ARGUMENTS

## Instructions

1. Parse the input above:
   - If `-e` is included → output in English
   - Otherwise → output in Japanese (default)
   - The remaining text after removing the option flag is the message body

2. **Transform the rough input into a formal business email:**
   - The input may be very casual, incomplete, or rough (e.g., "資料遅れる、すまん", "meeting delayed sorry")
   - Rewrite it completely as a polished, professional email
   - **IMPORTANT**: Always start with a greeting:
     - Japanese: 「お世話になっております。」
     - English: "Hope you're doing well." or "Thanks for your time."
   - Use appropriate honorifics and polite expressions (敬語 for Japanese)
   - Maintain a professional business tone throughout
   - Add appropriate closing remarks:
     - Japanese: 「何卒よろしくお願いいたします。」
     - English: "Best regards" or "Thank you for your understanding."

3. Output format:
   ```
   ---
   [Converted email body]
   ---
   ```

## Examples

### Japanese (default)
```
---
お世話になっております。

明日の会議につきまして、資料の準備が遅れており、
大変申し訳ございません。

本日中には完成させ、共有させていただく予定です。
何卒よろしくお願いいたします。
---
```

### English (-e)
```
---
Hope you're doing well.

Just wanted to let you know that the materials for
tomorrow's meeting are taking a bit longer than expected.
Sorry for the delay.

I'll make sure to have them completed and shared
by end of day.

Best regards
---
```
