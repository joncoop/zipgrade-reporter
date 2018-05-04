import json
import statistics
import docx

def get_import_path():
    '''
    Gets path to ZipGrade data file.
    
    Right now it's hard coded, but this could be obtained from GUI later.
    '''

    return "apcsa2015releasedexamExport.csv"
    #return "apcsa2014practiceexamExport.csv"

def get_export_path():
    '''
    Gets path to save report.
    
    Right now it's hard coded as an empty string so the report will be
    saved in the same file as the program is run.
    '''

    return ""

def get_export_filename():
    '''
    Gets path to save report.
    
    Right now it's hard coded, but this could be obtained from GUI later.
    '''

    return "sample_report.docx"

def csv_to_json(path):
    '''
    Reads ZipGrade csv export file and stores as JSON data
    '''
    
    with open(path, 'r') as f:
        contents = f.read().splitlines()

    fields = contents[0].split(',')

    records = []

    for line in contents[1:]:
        r = {}
        values = line.split(',')
        
        for field, value in zip(fields, values):
            r[field] = value

        records.append(r)

    #pretty = json.dumps(records, indent=4, sort_keys=False)
    #print(pretty)

    sort_by = lambda k: k['LastName'] + " " + k['FirstName']
    records = sorted(records, key=sort_by , reverse=False)
    
    return records

def get_raw_scores(records):
    scores = []

    for r in records:
        scores.append(int(float(r['EarnedPts'])))
        
    return scores
        
def get_percentages(records):
    scores = []

    for r in records:
        s = float(r['PercentCorrect'])
        scores.append(s)
        
    return scores

def make_meta_data(records, document):
    '''
    asdf
    '''

    r = records[0]
    date_created = r['QuizCreated']
    date_exported = r['DataExported']

    p = document.add_paragraph()
    p.add_run("Date Created: ")
    p.add_run(date_created + "\n")
    p.add_run("Date Exported: ")
    p.add_run(date_exported)

def make_summary_statistics(records, document):
    '''
    asdf
    '''
    r = records[0]
    possible_points = records[0]['PossiblePts']
    num_scores = len(records)
    
    scores = get_raw_scores(records)
    percentages = get_percentages(records)

    document.add_heading('Summary Statistics', 1)

    p = document.add_paragraph()
    p.add_run("Number of Scores: ")
    p.add_run(str(num_scores) + "\n")
    p.add_run("Points Possible: ")
    p.add_run(str(possible_points))

    mean_raw = round(statistics.mean(scores), 2)
    max_raw = max(scores)
    min_raw = min(scores)

    mean_percent = round(statistics.mean(percentages), 2)
    max_percent = max(percentages)
    min_percent = min(percentages)
    
    p = document.add_paragraph()
    p.add_run("Mean (raw/percent): ")
    p.add_run(str(mean_raw) + " / " + str(mean_percent) + "%\n")
    p.add_run("Max (raw/percent): ")
    p.add_run(str(max_raw) + " / " + str(max_percent) + "%\n")
    p.add_run("Min (raw/percent): ")
    p.add_run(str(min_raw) + " / " + str(min_percent) + "%")
    
def make_difficulty_analysis(records, document, hard_threshold, easy_threshold):
    '''
    Generates a list of questions ranked from most to least difficult
    based on the number of students that miss each question.
    '''
    r = records[0]
    num_questions = int(float(r['PossiblePts']))
    
    if 'Key100' in r:
        sheet_size = 100
    elif 'Key50' in r:
        sheet_size = 50
    else:
        sheet_size = 25

    misses = {}

    for r in records:
        for i in range(1, sheet_size + 1):
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

    document.add_heading('Difficulty Analysis', 1)

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
                               
def make_cover_page(records, document):
    '''
    Creates cover page with summary statistics.
    '''
    r = records[0]
    title = r['QuizName']

    document.add_heading('ZipGrade Score Report', 0)
    document.add_heading(title, 1)
    
    make_meta_data(records, document)
    make_summary_statistics(records, document)
    make_difficulty_analysis(records, document, 10, 2)
    
    document.add_page_break()

def make_score_summary(records, document):
    '''
    Creates summary of indidual student scores.
    '''

    document.add_heading('Individual Scores', 1)
    table = document.add_table(rows=1, cols=4)
    table.style = 'Medium Shading 1'
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Name'
    hdr_cells[1].text = 'Raw'
    hdr_cells[2].text = 'Possible'
    hdr_cells[3].text = 'Percent'

    for r in records:
        row_cells = table.add_row().cells
        row_cells[0].text = r['LastName'] + ", " + r['FirstName']
        row_cells[1].text = r['EarnedPts']
        row_cells[2].text = r['PossiblePts']
        row_cells[3].text = r['PercentCorrect'] + "%"
        
    document.add_page_break()

def make_individual_report(record, document):
    '''
    Creates report for a student
    '''
    
    title = record['QuizName']
    section = record['QuizClass']
    last = record['LastName']
    first = record['FirstName']
    zip_id = record['ZipGradeID']
    earned = record['EarnedPts']
    possible = record['PossiblePts']
    percent = record['PercentCorrect']

    if 'Key100' in record:
        sheet_size = 100
    elif 'Key50' in record:
        sheet_size = 50
    else:
        sheet_size = 25

    result = ""
    wrong = ""
    num_wrong = 0
    for i in range(1, sheet_size + 1):
        correct = record['Key' + str(i)]

        if len(correct) > 0:
            answer = record['Stu' + str(i)]

            if len(answer) == 0:
                answer = "-"
                
            num_wrong += 1
            wrong += "\t" + str(i) + ". " + answer

            if answer != correct:
                wrong += " (" + correct + ")"

                if i < 10:
                    wrong += "\t"
            else:
                wrong += "\t"
            
            if num_wrong % 5 == 0:
                wrong += "\n"
        
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.keep_together = True
        
    paragraph.add_run(last + ", " + first + "\n\n").bold = True
    paragraph.add_run("ID: " + zip_id + "\n")
    paragraph.add_run("Test: " + title + "\n")
    paragraph.add_run("Class: " + section + "\n")
    paragraph.add_run("Raw: " + earned + "/" + possible + "\n")
    paragraph.add_run("Percent: " +  percent + "%\n")
    paragraph.add_run("Response Summary: (Correct)\n")
    paragraph.add_run(wrong + "\n")
    paragraph.add_run("\n")
    
def generate_word_doc(records, save_path):
    '''
    Creates Word Doc with individual score reports.
    '''

    document = docx.Document()
    make_cover_page(records, document)
    make_score_summary(records, document)
    
    for r in records:
        make_individual_report(r, document)

    document.save(save_path)

# go!
if __name__ == "__main__":
    import_path = get_import_path()
    export_path = get_export_path()
    export_filename = get_export_filename()
    
    data = csv_to_json(import_path)
    save_path = export_path + export_filename
    generate_word_doc(data, save_path)
