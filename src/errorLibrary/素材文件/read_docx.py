# -*- coding: utf-8 -*-
import docx
import os

# Find file
base = r'C:\Users\Administrator\Desktop\clawTest\errorLibrary'
f = None
for root, dirs, files in os.walk(base):
    for file in files:
        if file.endswith('.docx'):
            f = os.path.join(root, file)
            break
    if f:
        break

print(f"File: {f}")

doc = docx.Document(f)
texts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
print(f"Total paragraphs: {len(texts)}")

# Save to temp file
output_path = os.path.join(os.path.dirname(f), 'temp_output.txt')
with open(output_path, 'w', encoding='utf-8') as fp:
    for t in texts:
        fp.write(t + '\n')

print(f"Saved to {output_path}")
