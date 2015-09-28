from member import Member
from docx import Document

csv_file_path = "/Users/Jorge/Projects/Ecusa/responses-sample.csv"
members_list = Member.load_from_csv(csv_file_path)
#
# html = Member.generate_html(members)
# html_file_path = "/Users/Jorge/Projects/Ecusa/html-sample.html"
# with open(html_file_path, "w") as f:
#     f.write(html)
# Member.generateDoc(members)

path = '/Users/Jorge/Desktop/doc2.docx'
f = open(path)
document = Document(f)
f.close()

mainCats = ["Cat1", "Cat2"]

for category in mainCats:
    document.add_heading(category, 1)

    style = "Medium Shading 1 Accent 3"
    table = document.add_table(rows=1, cols=1, style=style)
    table.autofit = False
    table.rows[0].cells[0].text = "Nombre"
    table.columns[0].width = 4839335

    # table.rows[0].cells[1].text = "Apellido1"
    # table.columns[1].width = 200
    # table.rows[0].cells[2].text = "Apellido2"
    # table.columns[2].width = 150
    for i in range(4):
        member = members_list[0]
        row = table.add_row()
        row.cells[0].text = member.first_name
        # row.cells[1].text = member.surname
        # row.cells[2].text = member.company



document.save(path)

#document.tables[0].rows[0].cells[0].width

