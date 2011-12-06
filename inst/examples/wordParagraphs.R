d = w[["Documents"]]$Add()

p = d[["Paragraphs"]]$Item(1)
r = p[["Range"]]
r[["Text"]] = "First paragraph's contents"

d[["Paragraphs"]]$Add()
p = d[["Paragraphs"]]$Item(d[["Paragraphs"]]$Count())
r = p[["Range"]]
r[["Text"]] = "Second paragraph's contents"


d[["Paragraphs"]]$Add()
p = d[["Paragraphs"]]$Item(d[["Paragraphs"]]$Count())
r = p[["Range"]]
r[["Text"]] = "Third paragraph's contents"