# card_extractor

Extracts debate card components from `.docx` files in sequence:
- Title (Heading 1, optional per card group)
- Tag (Heading 4)
- Citation (first non-empty paragraph after tag)
- Underlined quote text (collected after citation until next Heading 4)

Output is written without labels, in order:
title
tag
citation
underlined text
# card_extractor
# card_extractor
