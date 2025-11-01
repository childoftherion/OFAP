#!/bin/bash
for file in *.html; do
  statute_num=$(basename "$file" .html | tr '_' '.')
  md_file="${file%.html}.md"
  
  # Extract title
  title=$(sed -n 's/.*<span id="name">\([^<]*\)<\/span>.*/\1/p' "$file" | head -1)
  
  # Write header
  echo "# ORS $statute_num - $title" > "$md_file"
  echo "" >> "$md_file"
  echo "**Source**: https://oregon.public.law/statutes/$statute_num" >> "$md_file"
  echo "" >> "$md_file"
  echo "**Downloaded**: $(date '+%Y-%m-%d')" >> "$md_file"
  echo "" >> "$md_file"
  echo "---" >> "$md_file"
  echo "" >> "$md_file"
  
  # Extract section text - get content between <section> tags with class containing "outline"
  sed -n '/<section[^>]*class="[^"]*outline[^"]*"/,/<\/section>/p' "$file" | \
    sed 's/<[^>]*>//g' | \
    sed 's/^[[:space:]]*//' | \
    sed '/^$/d' >> "$md_file"
done
chmod +x extract_statutes.sh && ./extract_statutes.sh && echo "Done extracting statutes"