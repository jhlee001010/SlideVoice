import pymupdf as fitz

doc = fitz.open()
for i, title in enumerate(["첫 번째 슬라이드", "두 번째 슬라이드"], start=1):
    page = doc.new_page(width=1280, height=720)
    page.insert_text((100, 300), title, fontsize=48)
    page.insert_text((100, 400), f"page {i}", fontsize=24)
doc.save("sample/test.pdf")
print("wrote sample/test.pdf")
