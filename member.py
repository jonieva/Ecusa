import csv
from datetime import datetime

class Member(object):
    # Column indexes
    REGISTERED_DATE_HEADER = "Timestamp"
    # ID_IX = 0
    FIRST_NAME_HEADER = "Name"
    SURNAME_HEADER = "Last name"
    EMAIL_HEADER = "email"
    COMPANY_HEADER = "Affiliation / Institution / Company"
    POSITION_HEADER = "Job Title"
    MAIN_CATEGORY_HEADER = "What is your professional speciality?"
    KEYWORDS_HEADER = "Can you provide a few specific keywords to your professional activity?"
    LINKEDIN_HEADER = "LinkedIn profile"

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
    def load_from_csv(csv_file_path):
        """ Load a list of members from a csv file
        :param csv_file_path:
        :return:
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
                    print(member)
        return members
    
    @staticmethod
    def load_header(header_row):
        header = dict()
        header[Member.REGISTERED_DATE_HEADER] = header_row.index(Member.REGISTERED_DATE_HEADER)
        header[Member.FIRST_NAME_HEADER] = header_row.index(Member.FIRST_NAME_HEADER)
        header[Member.SURNAME_HEADER] = header_row.index(Member.SURNAME_HEADER)
        header[Member.EMAIL_HEADER] = header_row.index(Member.EMAIL_HEADER)
        header[Member.COMPANY_HEADER] = header_row.index(Member.COMPANY_HEADER)
        header[Member.POSITION_HEADER] = header_row.index(Member.POSITION_HEADER)
        header[Member.MAIN_CATEGORY_HEADER] = header_row.index(Member.MAIN_CATEGORY_HEADER)
        header[Member.KEYWORDS_HEADER] = header_row.index(Member.KEYWORDS_HEADER)
        header[Member.LINKEDIN_HEADER] = header_row.index(Member.LINKEDIN_HEADER)
        return header


    def load(self, row):
        if self.header is None:
            raise Exception("Header not loaded")
        self.registered_date = datetime.strptime(row[self.header[self.REGISTERED_DATE_HEADER]], "%m/%d/%Y %H:%M:%S" )
        self.first_name = row[self.header[self.FIRST_NAME_HEADER]]
        self.surname = row[self.header[self.SURNAME_HEADER]]
        self.email = row[self.header[self.EMAIL_HEADER]]
        self.company = row[self.header[self.COMPANY_HEADER]]
        self.position = row[self.header[self.POSITION_HEADER]]
        self.main_category = row[self.header[self.MAIN_CATEGORY_HEADER]]
        keywords = row[self.header[self.KEYWORDS_HEADER]].split("\n")
        for t in keywords:
            t = t.strip()
            if t != "":
                self.keywords.append(t)
        self.linkedin = row[self.header[self.LINKEDIN_HEADER]]






