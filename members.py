# -*- coding: utf-8 -*-
from member import Member
#from docx import Document


import csv
from datetime import datetime

import dominate
from dominate.tags import *

ENGLISH = 0
SPANISH = 1
FIELDS = {
    "first_name": ["First name", "Nombre"],
    "surname": ["Surname", "Apellidos"],
    "company": ["Company", "Empresa"],
    "position": ["Position", "Cargo"],
    "keywords": ["Specialities", "Áreas de especialidad"],
    "email": ["Email", "Email"],
    "linkedin": ["Web site / LinkedIn", "Página web / LinkedIn"],
}

CATEGORIES_LIST = ["Medicine", "Biology"]



def __getStyle__():
    style = \
        '''
        .CSSTableGenerator {
            margin:0px;padding:0px;
            width:100%;
            box-shadow: 10px 10px 5px #888888;
            border:1px solid #000000;

            -moz-border-radius-bottomleft:6px;
            -webkit-border-bottom-left-radius:6px;
            border-bottom-left-radius:6px;

            -moz-border-radius-bottomright:6px;
            -webkit-border-bottom-right-radius:6px;
            border-bottom-right-radius:6px;

            -moz-border-radius-topright:6px;
            -webkit-border-top-right-radius:6px;
            border-top-right-radius:6px;

            -moz-border-radius-topleft:6px;
            -webkit-border-top-left-radius:6px;
            border-top-left-radius:6px;
        }.CSSTableGenerator table{
            border-collapse: collapse;
                border-spacing: 0;
            width:100%;
            height:100%;
            margin:0px;padding:0px;
        }.CSSTableGenerator tr:last-child td:last-child {
            -moz-border-radius-bottomright:6px;
            -webkit-border-bottom-right-radius:6px;
            border-bottom-right-radius:6px;
        }
        .CSSTableGenerator table tr:first-child td:first-child {
            -moz-border-radius-topleft:6px;
            -webkit-border-top-left-radius:6px;
            border-top-left-radius:6px;
        }
        .CSSTableGenerator table tr:first-child td:last-child {
            -moz-border-radius-topright:6px;
            -webkit-border-top-right-radius:6px;
            border-top-right-radius:6px;
        }.CSSTableGenerator tr:last-child td:first-child{
            -moz-border-radius-bottomleft:6px;
            -webkit-border-bottom-left-radius:6px;
            border-bottom-left-radius:6px;
        }.CSSTableGenerator tr:hover td{
            background-color:#ffffff;


        }
        .CSSTableGenerator td{
            vertical-align:middle;
                background:-o-linear-gradient(bottom, #c6d66b 5%, #ffffff 100%);	background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #c6d66b), color-stop(1, #ffffff) );
            background:-moz-linear-gradient( center top, #c6d66b 5%, #ffffff 100% );
            filter:progid:DXImageTransform.Microsoft.gradient(startColorstr="#c6d66b", endColorstr="#ffffff");	background: -o-linear-gradient(top,#c6d66b,ffffff);

            background-color:#c6d66b;

            border:1px solid #000000;
            border-width:0px 1px 1px 0px;
            text-align:left;
            padding:7px;
            font-size:10px;
            font-family:Arial;
            font-weight:normal;
            color:#000000;
        }.CSSTableGenerator tr:last-child td{
            border-width:0px 1px 0px 0px;
        }.CSSTableGenerator tr td:last-child{
            border-width:0px 0px 1px 0px;
        }.CSSTableGenerator tr:last-child td:last-child{
            border-width:0px 0px 0px 0px;
        }
        .CSSTableGenerator tr th {
                background:-o-linear-gradient(bottom, #b8c958 5%, #b8c958 100%);	background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #b8c958), color-stop(1, #b8c958) );
            background:-moz-linear-gradient( center top, #b8c958 5%, #b8c958 100% );
            filter:progid:DXImageTransform.Microsoft.gradient(startColorstr="#b8c958", endColorstr="#b8c958");	background: -o-linear-gradient(top,#b8c958,b8c958);
            padding:7px;
            background-color:#b8c958;
            border:0px solid #000000;
            text-align:center;
            border-width:0px 0px 1px 1px;
            font-size:14px;
            font-family:Arial;
            font-weight:bold;
            color:#ffffff;
        }
        .CSSTableGenerator tr th:hover {
            /*background:-o-linear-gradient(bottom, #b8c958 5%, #b8c958 100%);	background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #b8c958), color-stop(1, #b8c958) );
            background:-moz-linear-gradient( center top, #b8c958 5%, #b8c958 100% );
            filter:progid:DXImageTransform.Microsoft.gradient(startColorstr="#b8c958", endColorstr="#b8c958");	background: -o-linear-gradient(top,#b8c958,b8c958);
            */
            background-color:#333333;
        }
        .CSSTableGenerator tr th {
            border-width:0px 0px 1px 0px;
        }
        .CSSTableGenerator tr th {
            border-width:0px 0px 1px 1px;
        }
        .CSSTableGenerator td:last-child {
            border-width:0px 0px 1px 1px;
        }

        .categoryTitle {
            color: #b8c958;
            font-size: 1.5em;
            font-weight: bold;
        }
        searchDiv {
            margin-top:30px
        }

        '''

    return style


def generate_html(members_list, language_code):
    """ Generate the html document for a list of members
    :param members_list: list of members objects
    :param language_code: 0 (ENGLISH) or 1 (SPANISH)
    :return: html
    """
    doc = dominate.document(title='ECUSA experts guide')

    with doc:
        with header():
            # CSS style
            style(__getStyle__(), type="text/css")

            # Javascript
            #script(type="text/javascript", src="http://listjs.com/no-cdn/list.js")
            script(type="text/javascript", src="http://code.jquery.com/jquery-1.11.3.min.js")
            script(type="text/javascript", src="sorttable.js")

            s = '''
                $(function(){
                    var options = {
                        valueNames: ['''
            for key in FIELDS.iterkeys():
                s+= "'{0}',".format(key)
            s = s[:-1]
            s+= "]};"


            # s+= '''
            #     $.extend($.expr[":"], {
            #         "containsNC": function(elem, i, match, array) {
            #             return (elem.textContent || elem.innerText || "").toLowerCase().indexOf((match[3] || "").toLowerCase()) >= 0;
            #         }
            #     });
            # '''

            s+= '''
                $.extend($.expr[":"], {
                    "containsNC": function(elem, i, match, array) {
                        var s = match[3].replace(";","|").toLowerCase();
                        var spl = "";
                        for (var i=0; i<spl.length; i++){
                            if (spl[i].trim() != "")
                                spl += spl[i].trim() + "|";
                        }
                        spl = spl.substring(0, spl.length -1);
                        re = new RegExp(spl);
                        return re.test((elem.textContent || elem.innerText || "").toLowerCase());
                    }
                });
            '''


            s+= '''
                    function refresh(){
                        var txt = $('#searchInput').val();

                        if (txt.length > 2){
                                res = $(".searchTable tbody tr:has(td:containsNC('" + txt + "'))");
                                $(res).show();
                                $(res).parent().parent().parent().show();

                                var res = $(".searchTable tbody tr:not(:has(td:containsNC('" + txt + "')))");
                                res.hide(300, function(){
                                    var divs = $(".searchDiv:has(.searchTable:not(:has(tbody tr:visible)))");
                                    divs.hide();
                                });
                        }
                        else {
                            $(".searchDiv .searchTable tbody tr").show();
                            $(".searchDiv").show();
                        }
                    }

                    $('#searchInput').keyup(function(e){
                        refresh();
                    });
                    $('#searchInput').change(function(){
                        refresh();
                    });
                });
                '''

            # Create the script without enconding
            scr = script(type="text/javascript")
            scr.add_raw_string(s)

        with body():
            with div():
                p("""La guia de expertos de ECUSA permite localizar rapidamente a profesionales que esten trabajando en
                    un determinado area de especialidad.""")
                p("Utiliza el campo de texto para buscar rapidamente en todos los campos de la tabla")
            with div(id="users"):
                with div("Search for any field"):
                    input(id="searchInput", cls="search", placeholder="Search")

                for category in CATEGORIES_LIST:
                    with div(id="div"+category, cls="searchDiv"):
                        p(category, cls="categoryTitle")
                        with table(cls="CSSTableGenerator searchTable sortable"):
                            # Header
                            with thead():
                                for key, values in FIELDS.iteritems():
                                    th(values[language_code], cls="sort").set_attribute("data-sort", key)
                            with tbody(cls="list"):
                                # Rows
                                for member in members_list:
                                    with tr():
                                        for key in FIELDS.iterkeys():
                                            s = eval("member.{0}".format(key))
                                            if isinstance(s, list):
                                                text = ", ".join(s)
                                            else:
                                                text = str(s)
                                            td(text, cls=key)

    html = doc.render()
    #print(doc)
    return html


def generate_Word_document(members_list):
    # document.save("/Users/Jorge/Desktop/mydoc.docx")
    path = '/Users/Jorge/Desktop/doc2.docx'
    f = open(path)
    document = Document(f)
    f.close()

    mainCats = ["Cat1", "Cat2"]
    #document.add_heading(u'Guía de expertos de ECUSA', 0)

    for category in mainCats:
        document.add_heading(category, 1)

        style = "Medium Shading 1 Accent 3"
        table = document.add_table(rows=1, cols=3, style=style)
        table.rows[0].cells[0].text = "Nombre"
        table.columns[0].width = 40
        table.rows[0].cells[1].text = "Apellido1"
        table.columns[1].width = 100
        table.rows[0].cells[2].text = "Apellido2"
        table.columns[2].width = 150
        for i in range(4):
            member = members_list[0]
            row = table.add_row()
            row.cells[0].text = member.first_name
            row.cells[1].text = member.surname
            row.cells[2].text = member.company
    document.save(path)





csv_file_path = "html-samples/responses-sample.csv"
members_list = Member.load_from_csv(csv_file_path)

html = generate_html(members_list, ENGLISH)
# print(html)
# html = completeHtml(html2)

html_file_path = "html-samples/html-sample6.html"
with open(html_file_path, "w") as f:
    f.write(html)


