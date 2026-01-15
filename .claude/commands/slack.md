---
model: haiku
---

# Slack Reply Generator

Convert rough messages into Slack-appropriate professional yet friendly messages.

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

2. Convert the message body for Slack:
   - **No greeting needed** (skip "お世話になっております" etc.)
   - Slightly more casual than email, but still business-appropriate
   - Concise and easy to read
   - Use line breaks effectively
   - Use emojis sparingly (only when appropriate)

3. Output format:
   - Output ONLY the converted message (no notes, no delimiters, no explanations)
   - ALWAYS copy the result to clipboard using Bash: `echo "..." | pbcopy`

## Examples

### Japanese (default)
```
明日の会議の件ですが、資料の準備が少し遅れています。
申し訳ありません。

本日中に共有しますので、少々お待ちください。
```

### English (-e)
```
Quick update on tomorrow's meeting -
the materials are taking a bit longer than expected.

I'll have them ready and shared by end of day.
Thanks for your patience!
```
