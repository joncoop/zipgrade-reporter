import docx
import json
import os
import statistics

from tkinter import *
from tkinter.filedialog import askopenfilename


class Scoresheet:
    def __init__(self, header_row, data_row, delimiter=","):
        """
        header_row is the top row of the ZipGrade csv export file.
        data_row is a single row containing one student's quiz data.
        """

        fields = header_row.split(delimiter)
        values = data_row.split(delimiter)

        def strip_quotes(s):
            if len(s) > 0 and s[0] == '"':
                s = s[1:]
            if len(s) > 0 and s[-1] == '"':
                s = s[:-1]

            return s.strip()
            
        values = [strip_quotes(v) for v in values]        
        data = {}
        
        for f, d in zip(fields, values):
            data[f] = d

        self.quiz_name = data['QuizName']
        self.class_name = data['QuizClass']
        self.first_name = data['FirstName']
        self.last_name = data['LastName']
        self.zip_id = data['ZipGradeID']
        self.external_id = data['ExternalID'] # unused
        self.earned_points = data['EarnedPts']
        self.possible_points = data['PossiblePts']
        self.percent_correct = data['PercentCorrect']
        self.date_created = data['QuizCreated']
        self.date_exported = data['DataExported']
        self.key_version = data['KeyVersion']
        self.num_questions = int((len(fields) - 12) / 4) # 12 metadata colums, 4 data cells per answer

        self.responses = []

        q = 1
        count = 0

        while count < self.num_questions:
            k = 'Key' + str(q)
            s = 'Stu' + str(q)
            
            if k in data:
                student_answer = data[s]
                correct_answer = data[k]
                r = {'question': q, 'answer': student_answer, 'correct': correct_answer}
                self.responses.append(r)
                count += 1

            q += 1


class Report:
    def __init__(self, scoresheets):
        self.scoresheets = scoresheets
        
    @property
    def versions(self):
        result = []

        for s in self.scoresheets:
            v = s.key_version

            if v not in result:
                result.append(v)

        result.sort()
        return result
    
    @property
    def classes(self):
        result = []

        for s in self.scoresheets:
            v = s.class_name

            if v not in result:
                result.append(v)

        result.sort()
        return result
    
    @property
    def raw_scores(self):
        result = []
        
        for s in self.scoresheets:
            n = float(s.earned_points)
            result.append(n)

        return result

    @property
    def percentages(self):
        result = []
        
        for s in self.scoresheets:
            n = float(s.percent_correct)
            result.append(n)

        return result
    
    def quartiles(self, num_list):
        nums = num_list.copy()
        nums.sort()

        mid1 = len(nums) // 2
        mid2 = mid1
        
        if len(nums) % 2 == 1:
            mid2 += 1
            
        q1 = round(statistics.median(nums[:mid1]), 2)
        q3 = round(statistics.median(nums[mid2:]), 2)

        return q1, q3 
    
    def add_report_title(self, document):
        sheet_1 = self.scoresheets[0]
        title = sheet_1.quiz_name

        document.add_heading('ZipGrade Score Report', 0)
        document.add_heading(title, 1)

    def add_meta_data(self, document):
        sheet_1 = self.scoresheets[0]

        p = document.add_paragraph()
        p.add_run("Date Created: ")
        p.add_run(sheet_1.date_created + "\n")
        p.add_run("Date Exported: ")
        p.add_run(sheet_1.date_exported)

        p = document.add_paragraph()
        p.add_run("Classes: " + "\n")
        for class_name in self.classes:
            p.add_run("  - " + class_name + "\n")
            
    def add_summary_statistics(self, document):
        sheet_1 = self.scoresheets[0]
        possible_points = sheet_1.possible_points
        num_scores = len(self.scoresheets)

        mean_raw = round(statistics.mean(self.raw_scores), 2)
        mean_pct = round(statistics.mean(self.percentages), 2)
        
        median_raw = round(statistics.median(self.percentages), 2)
        median_pct = round(statistics.median(self.percentages), 2)
        
        st_dev_raw = round(statistics.mean(self.raw_scores), 2)
        st_dev_pct = round(statistics.mean(self.percentages), 2)

        min_raw = round(min(self.raw_scores), 2)
        max_raw = round(max(self.raw_scores), 2)
        min_pct = round(min(self.percentages), 2)
        max_pct =  round(max(self.percentages), 2)
        
        q1_raw, q3_raw = self.quartiles(self.raw_scores)
        q1_pct, q3_pct = self.quartiles(self.percentages)
        
        document.add_heading('Summary Statistics', 1)

        p = document.add_paragraph()
        p.add_run("Number of Scores: ")
        p.add_run(str(num_scores) + "\n")
        p.add_run("Points Possible: ")
        p.add_run(str(possible_points))
        
        p = document.add_paragraph()
        p.add_run("Mean (raw/percent): ")
        p.add_run(str(mean_raw) + " / " + str(mean_pct) + "%\n")
        p.add_run("Standard Deviation (raw/percent): ")
        p.add_run(str(st_dev_raw) + " / " + str(st_dev_pct) + "%")
        
        p = document.add_paragraph()
        p.add_run("Max (raw/percent): ")
        p.add_run(str(max_raw) + " / " + str(max_pct) + "%\n")
        p.add_run("Q3 (raw/percent): ")
        p.add_run(str(q3_raw) + " / " + str(q3_pct) + "%\n")
        p.add_run("Median (raw/percent): ")
        p.add_run(str(median_raw) + " / " + str(median_pct) + "%\n")
        p.add_run("Q1 (raw/percent): ")
        p.add_run(str(q1_raw) + " / " + str(q1_pct) + "%\n")
        p.add_run("Min (raw/percent): ")
        p.add_run(str(min_raw) + " / " + str(min_pct) + "%")

    def add_individual_report(self, document, sheet):
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.keep_together = True
            
        paragraph.add_run(sheet.last_name + ", " + sheet.first_name + "\n\n").bold = True
        paragraph.add_run("ID: " + sheet.zip_id + "\n")
        paragraph.add_run("Test: " + sheet.quiz_name + "\n")
        if len(sheet.key_version) > 0:
            paragraph.add_run("Key: " + sheet.key_version + "\n")
        paragraph.add_run("Class: " + sheet.class_name + "\n")
        paragraph.add_run("Raw: " + sheet.earned_points + "/" + sheet.possible_points + "\n")
        paragraph.add_run("Percent: " +  sheet.percent_correct + "%\n")
        paragraph.add_run("Response Summary: (Correct)\n")

        count = 0
        paragraph.add_run('\t')

        ######### DO THIS AS A TABLE! NO TABS! #########
        for r in sheet.responses:
            q = str(r['question'])
            a = str(r['answer'])
            c = str(r['correct'])

            item = q + ". " + a
            if a != c:
                item += " (" + c + ")\t"
            else:
                item += '\t\t'
                
            paragraph.add_run(item)

            count += 1
            if count % 5 == 0:
               paragraph.add_run('\n\t')
            
        paragraph.add_run("\n")
        
    def save(self, document, path=None):
        document.save(path)
        
    def generate(self, save_path):
        document = docx.Document()

        # cover page
        self.add_report_title(document)
        self.add_meta_data(document)
        self.add_summary_statistics(document)
        document.add_page_break()
        # difficulty analysis
        
        # class reports

        # individual reports
        for s in self.scoresheets:
            self.add_individual_report(document, s)
            
        # save (or maybe return document and save there)
        self.save(document, save_path)
        
csv_file_path = '../APCSPTest12Export.csv'

with open(csv_file_path) as f:
    lines = f.readlines()

header = lines[0]
row = lines[1]

sheet_1 = Scoresheet(header, row)

print(sheet_1.quiz_name)

for r in sheet_1.responses:
    print(r)

all_sheets = []

for line in lines[1:]:
    sheet = Scoresheet(header, line)
    all_sheets.append(sheet)

r = Report(all_sheets)

print(r.versions)
print(r.classes)

save_path = "C:\\Users\\jccooper\\Desktop\\ZipGrade Reporter\\test.docx"
r.generate(save_path)





"""  
class App:
    def __init__(self, master):
        self.import_path = None
        self.export_path = None
        
        self.master = master
        self.gui_init()

    def gui_init(self):
        self.master.title("ZipGrade Reporter")
        
        select_button = Button(self.master, text="1. Select ZipGrade CSV Data", command=self.select_file)
        select_button.config(width=30)
        select_button.grid(row=0, column=0, padx=5, pady=5, sticky=(W))
        
        generate_button = Button(self.master, text="2. Generate Report", command=self.generate)
        generate_button.config(width=30)
        generate_button.grid(row=0, column=1, padx=5, pady=5, sticky=(E))

        instr1 = Label(self.master, text="The following data file will be used to generate your report...")
        instr1.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky=(W))

        self.import_lbl_text = StringVar()
        self.import_lbl_text.set("Waiting for file selection...")
        import_lbl = Label(self.master, textvariable=self.import_lbl_text)
        import_lbl.grid(row=4, column=0, columnspan=2, padx=20, pady=5, sticky=(W))

        instr2 = Label(self.master, text="Report will be created in...")
        instr2.grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky=(W))
        
        self.export_lbl_text = StringVar()
        self.export_lbl_text.set("...")
        export_lbl = Label(self.master, textvariable=self.export_lbl_text)
        export_lbl.grid(row=6, column=0, columnspan=2, padx=20, pady=5, sticky=(W))
        
        self.status_lbl_text = StringVar()
        status_lbl = Label(self.master, textvariable=self.status_lbl_text)
        status_lbl.grid(row=7, column=0, columnspan=2, padx=5, pady=5, sticky=(W))

    def select_file(self):
        '''
        Gets path to ZipGrade data file and sets export path to same directory.
        '''
        
        self.import_path = askopenfilename()
        self.export_path = os.path.dirname(self.import_path)
        
        self.import_lbl_text.set(self.import_path)
        self.export_lbl_text.set(self.export_path)
        
    def change_export_path(self):
        '''
        Maybe put a button next to the export field so that it can be changed if desired
        '''
        pass
    
    def get_export_filename(self, records):
        '''
        Gets path to save report. The report file name is simply the quiz name
        and the export date. If no quiz name exists, then the name will default
        to grade_report
        
        ZipGrade date format: May 02 2018 02:14 PM
        '''
        
        months = {"Jan": "01",
                  "Feb": "02",
                  "Mar": "03",
                  "Apr": "04",
                  "May": "05",
                  "Jun": "06",
                  "Jul": "07",
                  "Aug": "08",
                  "Sep": "09",
                  "Oct": "10",
                  "Nov": "11",
                  "Dec": "12"}
                  
        r = records[0]
        
        title = r['QuizName'].strip()
        if len(title) == 0:
            title = "grade_report"

        '''
        # can't make this part of file name, a single report might contain multiple classes
        section = r['QuizClass'].strip()
        if len(section) == 0:
            section = ""
        '''
        
        date = r['DataExported'].split(" ")
        yyyy = date[2]
        mm = months[date[0]]
        dd = date[1]
            
        if int(dd) < 10:
            dd = "0" + dd

        #hr, mi = date[3].split(":")
        #period = date[4]

        temp = title + "_" + "_" + yyyy + mm + dd
        filename = ""
        underscore = True

        for c in temp:
            if c.isalnum():
                filename += c
                underscore = False
            elif c == "_" and underscore == False:
                filename += c
                underscore = True
            elif underscore == False:
                filename += "_"
                underscore = True
            
        return filename  + ".docx"
        
    def make_difficulty_analysis(self, records, document):
        '''
        Generates a list of questions ranked from most to least difficult
        based on the number of students that miss each question.
        '''

        document.add_heading('Difficulty Analysis', 1)
        
        versions = self.get_versions(records)

        for v in versions:
            misses = {}
            r = records[0]
            num_questions = int((len(r) - 12) / 4) # 12 metadata colums, 4 data cells per answer
        
            if len(versions) > 1:
                document.add_heading('Key version: ' + v, 2)
                
            for r in records:
                if r['KeyVersion'] == v:
                    for i in range(1, num_questions + 1):
                        correct = r['Key' + str(i)]
                        answer = r['Stu' + str(i)]
                        
                        if len(correct) > 0:
                            if i not in misses:
                                misses[i] = 0
                                
                            if answer != correct:
                                misses[i] += 1

            difficulty = []
            for k, v in misses.items():
                p = round(v / num_questions * 100, 1)
                difficulty.append((k, v, p))

            sort_by = lambda k: k[1]
            difficulty = sorted(difficulty, key=sort_by , reverse=True)

            # better idea to use standard deviations for analysis?
            if len(difficulty) > 10:
                hard_threshold = difficulty[4][2]
                easy_threshold = difficulty[-3][2]
            
                paragraph = document.add_paragraph("Most difficult Questions (at least " + str(hard_threshold) + "% missed)\n")
                for d in difficulty:
                    if (d[2] >= hard_threshold):
                        q = str(d[0])
                        n = str(d[1])
                        p = str(d[2])
                        paragraph.add_run("\tq=" + q + ", n=" + n + ", %=" + p + "\n")

                paragraph = document.add_paragraph("Easiest Questions (no more than " + str(easy_threshold) + "% missed)\n")
                for d in difficulty:
                    if (d[2] <= easy_threshold):
                        q = str(d[0])
                        n = str(d[1])
                        p = str(d[2])
                        paragraph.add_run("\tq=" + q + ", n=" + n + ", %=" + p + "\n")
            else:
                for d in difficulty:
                    q = str(d[0])
                    n = str(d[1])
                    p = str(d[2])
                    paragraph.add_run("\tq=" + q + ", n=" + n + ", %=" + p + "\n")
                    
        document.add_page_break()
                                           
        
    
    def make_score_summary(self, records, document):
        '''
        Creates summary of indidual student scores.
        '''
        classes = self.get_classes(records)
        
        for c in classes:
            if c != '':
                document.add_heading('Class scores for ' + c, 1)
            else:
                document.add_heading('Class scores', 1)
            
            table = document.add_table(rows=1, cols=4)
            table.style = 'Medium Shading 1'
            
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Name'
            hdr_cells[1].text = 'Raw'
            hdr_cells[2].text = 'Possible'
            hdr_cells[3].text = 'Percent'

            for r in records:
                if c == r['QuizClass']:
                    row_cells = table.add_row().cells
                    row_cells[0].text = r['LastName'] + ", " + r['FirstName']
                    row_cells[1].text = r['EarnedPts']
                    row_cells[2].text = r['PossiblePts']
                    row_cells[3].text = r['PercentCorrect'] + "%"
                
            document.add_page_break()



    def save(self, document, records):
        '''
        Creates Word Doc with individual score reports.
        '''                            
        try:
            self.save_path = self.export_path + "/" + self.get_export_filename(records)
            document.save(self.save_path)
            self.status_lbl_text.set("Your report is ready!")
        except:
            self.status_lbl_text.set("Unable to save report. Check file and disk permissions.")

    def make_student_report_cover(self, class_name, document):
        paragraph = document.add_paragraph()
        paragraph.add_run('\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n')
        heading = document.add_heading('Individual student Reports for\n' + class_name, 1)
        heading.alignment = 1
        document.add_page_break()
        
    def generate(self):
        '''
        Creates Word Doc with individual score reports.
        '''
        generated = False
        
        if self.import_path != None:
            try:
                records = self.csv_to_json(self.import_path)

                document = docx.Document()
                self.make_cover_page(records, document)
                self.make_difficulty_analysis(records, document)
                self.make_score_summary(records, document)
                
                classes = self.get_classes(records)
                
                for i, c in enumerate(classes):
                    if c != '':
                        self.make_student_report_cover(c, document)
                        
                    for r in records:
                        if c == r['QuizClass']:
                            self.make_individual_report(r, document)

                    if i < len(classes) - 1:
                        document.add_page_break()

                generated = True
            except Exception as inst:
                print(inst)
                self.status_lbl_text.set("Something went wrong. Be sure your CSV data file is valid.")

            if generated:
                self.save(document, records)
        else:
            self.status_lbl_text.set("You must select a file first!")


root = Tk()
#root.iconbitmap('assets/my_icon.ico')
my_gui = App(root)
root.mainloop()
"""
