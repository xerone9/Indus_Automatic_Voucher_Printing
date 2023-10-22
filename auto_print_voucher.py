import PyPDF2
import win32print
import win32api
import requests


def get_data():
    url = "https://script.google.com/macros/s/AKfycbztIlguZFMWj3fNwIwUpIG7XDyBxB2u2b4Li6HZWwCIsPHbs3eMxJrKM_BfaAXexQC8/exec"  # Replace with the actual URL

    try:
        response = requests.get(url)

        if response.status_code == 200:
            data = response.json()
            return data
        else:
            return "Request failed with status code:", response.status_code

    except requests.exceptions.RequestException as e:
        return "Request failed:", e


def parse_data():
    data = {}
    for item in get_data()['data']:
        student_id_barcode = item['student_id_barcode']
        student_id = item['student_id']
        student_voucher = item['student_voucher']

        data[student_id_barcode] = [student_id, student_voucher]

    URL = get_data()['URL']
    return URL, data


def download_voucher(id, url):
    output_file_path = id + ".pdf"
    response = requests.get(url)

    if response.status_code == 200:
        with open(output_file_path, 'wb') as pdf_file:
            pdf_file.write(response.content)
        return output_file_path
    else:
        return None


def rotate_first_page(input_pdf, output_pdf):
    pdf_reader = PyPDF2.PdfReader(open(input_pdf, 'rb'))
    first_page = pdf_reader.pages[0]
    first_page.rotate(90)

    pdf_writer = PyPDF2.PdfWriter()
    pdf_writer.add_page(first_page)

    with open(output_pdf, 'wb') as output_file:
        pdf_writer.write(output_file)


def print_voucher(voucher):
    GHOSTSCRIPT_PATH = "C:/Program Files/gs/gs10.02.0/bin/gswin64.exe"
    GSPRINT_PATH = "C:/Program Files/Ghostgum/gsview/gsprint.exe"
    currentprinter = win32print.GetDefaultPrinter()

    win32api.ShellExecute(0, 'open', GSPRINT_PATH,
                          '-ghostscript "' + GHOSTSCRIPT_PATH + '" -printer "' + currentprinter + '" "' + voucher + '"',
                          '.', 0)


def main():
    print("Fetching Data...\n")

    URL = parse_data()[0]
    data = parse_data()[1]

    print("System Ready...\n")

    while True:
        barcode = input("Punch Card: ")
        if data.get(int(barcode)):
            generate_url = str(URL).replace("intercept1", data[int(barcode)][0]).replace("intercept2", str(data[int(barcode)][1]))
            input_pdf_path = download_voucher(data[int(barcode)][0], generate_url)

            output_pdf_path = input_pdf_path.split(".")[0] + ' rotated.pdf'
            rotate_first_page(input_pdf_path, output_pdf_path)

            print_voucher(output_pdf_path)

        else:
            print("Voucher Not Found")


if __name__ == "__main__":
    main()