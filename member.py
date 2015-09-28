# coding=UTF-8

import csv
from datetime import datetime

import dominate
from dominate.tags import *

from docx import Document

class Member(object):
    # Column indexes
    __REGISTERED_DATE_HEADER__ = "Timestamp"
    # ID_IX = 0
    __FIRST_NAME_HEADER__ = "Name"
    __SURNAME_HEADER__ = "Last name"
    __EMAIL_HEADER__ = "email"
    __COMPANY_HEADER__ = "Affiliation / Institution / Company"
    __POSITION_HEADER__ = "Job Title"
    __MAIN_CATEGORY_HEADER__ = "What is your professional speciality?"
    __KEYWORDS_HEADER__ = "Can you provide a few specific keywords to your professional activity?"
    __LINKEDIN_HEADER__ = "LinkedIn profile"

    def __init__(self):
        self.header = {}

        self.id = 0
        self.registered_date = None
        self.first_name = ""
        self.surname = ""
        self.email = ""
        self.company = ""
        self.position = ""
        self.main_category = ""
        self.keywords = []
        self.linkedin = ""


    def __str__(self):
        return  "** Member {0}\n" \
                "Registered date: {1}\n" \
                "First name: {2}\n"\
                "Surname: {3}\n"\
                "Email: {4}\n"\
                "Company: {5}\n"\
                "Position: {6}\n"\
                "Main category: {7}\n"\
                "Keywords: {8}\n"\
                "LinkedIn: {9}" \
                .format(
                    self.id,
                    self.registered_date,
                    self.first_name,
                    self.surname,
                    self.email,
                    self.company,
                    self.position,
                    self.main_category,
                    "; ".join(self.keywords),
                    self.linkedin
                )

    @staticmethod
    def __getCategories__():
        return ["Cat1", "Cat2"]


    @staticmethod
    def load_from_csv(csv_file_path):
        """ Load a list of members from a csv file
        :param csv_file_path: path to the csv file
        :return: list of members
        """
        members = []
        with open(csv_file_path, 'rb') as f:
            reader = csv.reader(f)
            header = None
            for row in reader:
                if header is None:
                    header = Member.load_header(row)
                else:
                    member = Member()
                    member.header = header
                    member.load(row)
                    members.append(member)
                    #print(member)
        return members
    
    @staticmethod
    def load_header(header_row):
        """ Build the header dictionary that will be used to load a member's data from a row.
        For every property, it will store the CSV file header key and the corresponding int index in the file
        :param header_row: row that will contain the headers of the csv file
        :return: dictionary with Key-Colum_index
        """
        header = dict()
        header[Member.__REGISTERED_DATE_HEADER__] = header_row.index(Member.__REGISTERED_DATE_HEADER__)
        header[Member.__FIRST_NAME_HEADER__] = header_row.index(Member.__FIRST_NAME_HEADER__)
        header[Member.__SURNAME_HEADER__] = header_row.index(Member.__SURNAME_HEADER__)
        header[Member.__EMAIL_HEADER__] = header_row.index(Member.__EMAIL_HEADER__)
        header[Member.__COMPANY_HEADER__] = header_row.index(Member.__COMPANY_HEADER__)
        header[Member.__POSITION_HEADER__] = header_row.index(Member.__POSITION_HEADER__)
        header[Member.__MAIN_CATEGORY_HEADER__] = header_row.index(Member.__MAIN_CATEGORY_HEADER__)
        header[Member.__KEYWORDS_HEADER__] = header_row.index(Member.__KEYWORDS_HEADER__)
        header[Member.__LINKEDIN_HEADER__] = header_row.index(Member.__LINKEDIN_HEADER__)
        return header


    def load(self, row):
        """ Load the information of a member from a csv row
        :param row: csv row (as a list)
        """
        if self.header is None:
            raise Exception("Header not loaded")
        self.registered_date = datetime.strptime(row[self.header[self.__REGISTERED_DATE_HEADER__]], "%m/%d/%Y %H:%M:%S" )
        self.first_name = row[self.header[self.__FIRST_NAME_HEADER__]]
        self.surname = row[self.header[self.__SURNAME_HEADER__]]
        self.email = row[self.header[self.__EMAIL_HEADER__]]
        self.company = row[self.header[self.__COMPANY_HEADER__]]
        self.position = row[self.header[self.__POSITION_HEADER__]]
        self.main_category = row[self.header[self.__MAIN_CATEGORY_HEADER__]]
        keywords = row[self.header[self.__KEYWORDS_HEADER__]].split("\n")
        for t in keywords:
            t = t.strip()
            if t != "":
                self.keywords.append(t)
        self.linkedin = row[self.header[self.__LINKEDIN_HEADER__]]

    @staticmethod
    def generate_html(members_list):
        """ Generate the html document for a list of members
        :param members_list: list of members objects
        :return: html
        """
        doc = dominate.document(title='ECUSA experts guide')

        with doc:
            with div(id="members").add(table()):
                # Header
                th("Date")
                th("First name")
                th("Surname")
                th("Email")

                # Rows
                for member in members_list:
                    # member = Member()
                    with tr():
                        td(str(member.registered_date), cls="myClass")
                        td(member.first_name, cls="myClass")
                        td(member.surname, cls="myClass")
                        td(member.email, cls="myClass")
        html = doc.render()
        #print(doc)
        return html


    @staticmethod
    def generateDoc(members_list):
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

