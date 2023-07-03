import PyPDF2
import re
import pandas as pd


def extract_text_from_pdf(pdf_file: str, start_page: int) -> str:
    text = ""
    with open(pdf_file, 'rb') as pdf:
        reader = PyPDF2.PdfReader(pdf, strict=False)
        for page_number in range(start_page - 1, len(reader.pages)):
            page = reader.pages[page_number]
            text += page.extract_text()
    return text


def remove_square_brackets(text: str) -> str:
    return re.sub(r'\[.*?\]', '', text)


def remove_rnds(text: str) -> str:
    return re.sub(r'RNDS.*', '', text)


def extract_can_lines(text: str):
    lines = text.split('\n')
    data = {'stakeholder_ids': [], 'filtered_data': []}
    current_id = ""
    current_data = ""
    for line in lines:
        if line.startswith('[CAN') or line.startswith('[CANGM'):
            if current_id:
                filtered_data = remove_square_brackets(f'{current_id} {current_data}').strip()
                data['filtered_data'].append(remove_rnds(filtered_data))
            data['stakeholder_ids'].append(re.findall(r'\[(.*?)\]', line)[0].strip())
            current_id = line
            current_data = ""
        elif current_id and 'Figure' not in line:
            current_data += ' ' + line

    if current_id:
        filtered_data = remove_square_brackets(f'{current_id} {current_data}').strip()
        data['filtered_data'].append(remove_rnds(filtered_data))

    df = pd.DataFrame(data)
    df.to_excel('Test Data1.xlsx', index=False, startrow=0, startcol=0)
    print("Data exported to 'Test Data1.xlsx'")


if __name__ == '__main__':
    pdf_text = extract_text_from_pdf('test.pdf', start_page=10)
    extract_can_lines(pdf_text)
