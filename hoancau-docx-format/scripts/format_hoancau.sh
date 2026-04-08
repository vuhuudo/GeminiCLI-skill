#!/usr/bin/env bash

# This script applies the HoanCauGroup 2025 formatting standards to a .docx file using officecli.

OFFICECLI="/home/johndoe/.local/bin/officecli"
MKTEMP="/usr/bin/mktemp"
RM="/usr/bin/rm"

FILE=$1

if [ -z "$FILE" ]; then
    echo "Usage: $0 <file.docx>"
    exit 1
fi

<<<<<<< HEAD
# 1. Page Setup
=======
# 1. Page Setup (A4: 21.0 x 29.7 cm)
# Margins: Top 2.0cm, Bottom 2.0cm, Left 3.0cm, Right 1.5cm
>>>>>>> f112a90 (feat(hoancau): add HoanCau DOCX Format skill (standard 2026))
$OFFICECLI set "$FILE" / \
    --prop pageWidth=11906 \
    --prop pageHeight=16838 \
    --prop marginTop=1134 \
    --prop marginBottom=1134 \
    --prop marginLeft=1701 \
    --prop marginRight=850

# 2. General Paragraph Formatting
<<<<<<< HEAD
=======
# Body: 13pt, Justified, 1.3x line spacing, first line indent 1.25cm
>>>>>>> f112a90 (feat(hoancau): add HoanCau DOCX Format skill (standard 2026))
PARA_PATHS=$($OFFICECLI query "$FILE" 'paragraph' --json | grep -oE '"path": "/body/p[^"]+"' | cut -d'"' -f4)

BATCH_FILE=$($MKTEMP)
echo "[" > "$BATCH_FILE"
FIRST=true
for P_PATH in $PARA_PATHS; do
<<<<<<< HEAD
=======
    # Skip paragraphs inside tables as they usually have special formatting
>>>>>>> f112a90 (feat(hoancau): add HoanCau DOCX Format skill (standard 2026))
    if [[ "$P_PATH" == *"/tbl["* ]]; then
        continue
    fi
    
<<<<<<< HEAD
=======
    # Get paragraph text to check for headings
    P_TEXT=$($OFFICECLI get "$FILE" "$P_PATH" --json | grep -oE '"text": "[^"]+"' | cut -d'"' -f4)
    
    # Default props for body text
    PROPS="\"font\":\"Times New Roman\",\"size\":\"13pt\",\"alignment\":\"justify\",\"lineSpacing\":\"1.3x\",\"spaceBefore\":\"3pt\",\"spaceAfter\":\"6pt\",\"firstLineIndent\":\"1.25cm\",\"widowControl\":false"
    
    # Simple heading detection (e.g., "1. ", "2. ", "3.1. ")
    if [[ "$P_TEXT" =~ ^[0-9]+(\.[0-9]+)*\. ]]; then
        PROPS="\"font\":\"Times New Roman\",\"size\":\"13pt\",\"alignment\":\"left\",\"bold\":true,\"lineSpacing\":\"1.3x\",\"spaceBefore\":\"6pt\",\"spaceAfter\":\"6pt\",\"firstLineIndent\":\"0cm\""
    fi

    # Special case for "TỜ TRÌNH"
    if [[ "$P_TEXT" == "TỜ TRÌNH" ]]; then
        PROPS="\"font\":\"Times New Roman\",\"size\":\"16pt\",\"alignment\":\"center\",\"bold\":true,\"lineSpacing\":\"1.3x\",\"spaceBefore\":\"12pt\",\"spaceAfter\":\"12pt\",\"firstLineIndent\":\"0cm\""
    fi
    
    # Special case for "V/v:"
    if [[ "$P_TEXT" =~ ^V/v: ]]; then
        PROPS="\"font\":\"Times New Roman\",\"size\":\"13pt\",\"alignment\":\"center\",\"bold\":true,\"italic\":true,\"lineSpacing\":\"1.3x\",\"spaceBefore\":\"0pt\",\"spaceAfter\":\"12pt\",\"firstLineIndent\":\"0cm\""
    fi

    # Special case for "Kính gửi:"
    if [[ "$P_TEXT" =~ ^"Kính gửi:" ]]; then
        PROPS="\"font\":\"Times New Roman\",\"size\":\"13pt\",\"alignment\":\"left\",\"bold\":true,\"lineSpacing\":\"1.3x\",\"spaceBefore\":\"12pt\",\"spaceAfter\":\"6pt\",\"firstLineIndent\":\"0cm\""
    fi

>>>>>>> f112a90 (feat(hoancau): add HoanCau DOCX Format skill (standard 2026))
    if [ "$FIRST" = true ]; then
        FIRST=false
    else
        echo "," >> "$BATCH_FILE"
    fi
<<<<<<< HEAD
    echo "{\"command\":\"set\",\"path\":\"$P_PATH\",\"props\":{\"font\":\"Times New Roman\",\"size\":\"13pt\",\"alignment\":\"justify\",\"lineSpacing\":\"1.3x\",\"spaceBefore\":\"3pt\",\"spaceAfter\":\"6pt\",\"firstLineIndent\":\"1.25cm\",\"widowControl\":false}}" >> "$BATCH_FILE"
=======
    echo "{\"command\":\"set\",\"path\":\"$P_PATH\",\"props\":{$PROPS}}" >> "$BATCH_FILE"
>>>>>>> f112a90 (feat(hoancau): add HoanCau DOCX Format skill (standard 2026))
done
echo "]" >> "$BATCH_FILE"

$OFFICECLI open "$FILE"
$OFFICECLI batch "$FILE" < "$BATCH_FILE"
$OFFICECLI close "$FILE"

# 3. Footer
FOOTER_PATH=$($OFFICECLI add "$FILE" / --type footer --prop text=" " --prop alignment=center --prop size=12pt --prop font="Times New Roman" | grep -oE '/footer\[[0-9]+\]')

if [ -z "$FOOTER_PATH" ]; then
    FOOTER_PATH="/footer[1]"
fi

$OFFICECLI raw-set "$FILE" "$FOOTER_PATH" \
  --xpath "//w:p" \
  --action append \
  --xml '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>'

$RM "$BATCH_FILE"

echo "Formatting completed for $FILE (HoanCauGroup 2025)."
