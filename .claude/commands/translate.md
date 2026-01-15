# English Translation

Translate Japanese messages into natural English while preserving the original meaning.

## Tone Options
- (default): Polite, slightly formal English
- `-c`: Casual, relaxed English

## Input
$ARGUMENTS

## Instructions

1. Parse the input above:
   - If `-c` is included → use casual/relaxed tone
   - Otherwise → use polite/slightly formal tone (default)
   - The remaining text after removing the option flag is the message to translate

2. Analyze the input Japanese text:
   - Understand the meaning and intent
   - Grasp the nuances

3. Translate into English:

   **Default (polite):**
   - Use polite and professional expressions
   - Use courteous phrases (e.g., "I would appreciate...", "Could you please...")
   - Slightly formal but natural

   **Casual (-c):**
   - Use relaxed, everyday expressions
   - Shorter sentences, contractions OK (e.g., "I'll", "Can you...")
   - Friendly and approachable tone

3. Output format:
   ```
   ---
   [English translation]
   ---

   **Notes** (only if needed)
   - Include translation tips or alternative expressions
   ```

## Examples

### Default (polite)
Input: `お忙しいところ恐れ入りますが、ご確認いただけますでしょうか。`

```
---
I understand you are busy, but I would greatly appreciate it if you could review this at your earliest convenience.
---
```

### Casual (-c)
Input: `-c お忙しいところ恐れ入りますが、ご確認いただけますでしょうか。`

```
---
Sorry to bother you, but can you take a look at this when you get a chance?
---
```
