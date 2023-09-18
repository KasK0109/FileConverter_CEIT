import os
import sys
import openpyxl
import requests
from bs4 import BeautifulSoup
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Font, Border, Side, PatternFill, NamedStyle


# Funktion til at hente data fra tekst fil
def extract_data(file_path):
    data = {}
    with (open(file_path, 'r', encoding='ansi') as file):
        lines = file.readlines()
        data['BrugerNavn'] = lines[0].strip()  # Hent bruger navn fra første linje
        free_disc_space = lines[-1].strip()
        if free_disc_space.startswith('2 Verzeichnis(se)'):
            data['FriDiskPlads>\C:'] = free_disc_space[19:24]
        else:
            data['FriDiskPlads>\C:'] = free_disc_space[10:15]  # Hent fri disk plads

        for line in lines:
            if line.startswith('OS Version:') or line.startswith('Version du systŠme'):
                winVersion = line.strip().split(':')[1].strip().split()  # Hent Windows Version og split for join
                data['WindowsVersion'] = ' '.join([winVersion[0], winVersion[2], winVersion[3]])
            if line.startswith('Betriebssystemversion'):
                winVersion = line.strip().split(':')[1].strip().split()
                data['WindowsVersion'] = ' '.join([winVersion[0], winVersion[3], winVersion[4]])
            if line.startswith('Host Name:') or line.startswith('Hostname:') or line.startswith('Nom de l'):
                data['HostName'] = line.strip().split(':')[1].strip()  # Hent Host Navn
            if line.startswith('OS Name:') or line.startswith('Betriebssystemname:') or line.startswith(
                    'Nom du systŠme'):
                data['OSName'] = line.strip().split(':')[1].strip()  # Hent OS Navn
            if line.startswith('System Locale') or line.startswith('Systemgebietsschema') or line.startswith(
                    'Option r‚gionale du systŠme'):
                data['SysLang'] = line.strip().split(':')[1].strip()

    return data


def convert_french_format_to_number(french_format):
    # Fjern "octets libres" fra strengen og erstatter "ÿ" med intet
    cleaned_format = french_format.replace("ÿ", ",")

    return cleaned_format


def get_latest_windows_versions(url):  # Find de seneste windows versioner

    # Hent content
    response = requests.get(url)
    html_content = response.content

    soup = BeautifulSoup(html_content, "html.parser")

    # Find elmenter med seneste version og build
    # Find elmenter med seneste version og build
    target_element = ""

    if url.__contains__("windows11"):
        target_element = soup.select(
            "html.hasSidebar.hasPageActions.hasBreadcrumb.conceptual.has-default-focus.theme-light"
            " body div.mainContainer"
            ".uhf-container.has-default-focus div.columns.has-large-gaps.is-gapless-mobile "
            "section.primary-holder.column.is-two-thirds-tablet.is-three-quarters-desktop div.columns."
            "is-gapless-mobile.has-large-gaps"
            " div#main-column.column.is-full.is-8-desktop main#main div.content div#winrelinfo_container strong")

    elif url == "https://learn.microsoft.com/en-us/windows/release-health/release-information":
        target_element = soup.select(
            "html.hasSidebar.hasPageActions.hasBreadcrumb.conceptual.has-default-focus.theme-light body "
            "div.mainContainer.uhf-container.has-default-focus div.columns.has-large-gaps.is-gapless-mobile "
            "section.primary-holder.column.is-two-thirds-tablet.is-three-quarters-desktop"
            " div.columns.is-gapless-mobile.has-large-gaps "
            "div#main-column.column.is-full.is-8-desktop main#main div.content div#winrelinfo_container strong")

    if target_element:
        # Extract the content of the target element
        content = target_element[0].get_text().strip()
        return content
    else:
        print("Target element not found.")


try:
    win11_version = get_latest_windows_versions(
        "https://learn.microsoft.com/en-us/windows/release-health/windows11-release-information"
    ).split(" ")[4].split(")")[0]
    win10_version = get_latest_windows_versions(
        "https://learn.microsoft.com/en-us/windows/release-health/release-information"
    ).split(" ")[4].split(")")[0]

except Exception as e:
    print("Der er sket en fejl i indhentningen af de seneste Windows Versioner")
    print("Check on der er forbindelse til internettet!")
    print(f"Fejlen: {e}")


def get_executable_directory():  # Find den korrekte PATH til mappen med tekst filer
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


folder_path = get_executable_directory()
print(folder_path)

# Start excel bog
workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet.title = 'Data'

# Border style
medium_border = Border(
    left=Side(style='medium'),
    right=Side(style='medium'),
    top=Side(style='medium'),
    bottom=Side(style='medium')
)

# Fed tekst
bold_font = Font(bold=True)

# Skriv data navne
worksheet.append(['Bruger Navn', 'Host Navn', 'OS Navn', 'Windows Version', 'Fri disk plads på C:'])

number_format_style = NamedStyle(name='number_format_style', number_format='##0,0')
target_column = "E"
worksheet[f'{target_column}1'].style = number_format_style

# Løkke til at løbe igennem tekst filer i mappe og hente data
try:
    for filename in os.listdir(folder_path):
        if filename.endswith('.txt'):
            file_path = os.path.join(folder_path, filename)
            data = extract_data(file_path)

            # Skriv data til excel fil
            row = [data.get('BrugerNavn', ''),
                   data.get('HostName', ''),
                   data.get('OSName', ''),
                   data.get('WindowsVersion', ''),
                   ]

            if data.get('SysLang').startswith('fr'):
                fri_disk_plads = convert_french_format_to_number(data.get('FriDiskPlads>\C:'))
            else:
                fri_disk_plads = data.get('FriDiskPlads>\C:', '').replace('.', ',')

            row.append(float(fri_disk_plads.replace(',', '.')))

            worksheet.append(row)

    header_row = worksheet[1]
    for cell in header_row:
        cell.font = Font(bold=True)
        cell.border = Border(bottom=Side(style='thick'))

    for column in worksheet.columns:
        max_length = max(len(str(cell.value)) for cell in column)
        adjusted_width = (max_length + 2) * 1.2  # Juster 1.2
        worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

    # Define fill colors for cells
    green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    yellow_fill = PatternFill(start_color="E4CD00", end_color="E4CD00", fill_type="solid")

    # Iterate through rows (excluding the header row)
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
        os_name = row[2].value
        win_version_row = row[3].value
        win_version = win_version_row[-5:]

        if os_name == "Microsoft Windows 11 Pro":
            if win_version == win11_version:
                row[3].fill = green_fill
            else:
                row[3].fill = red_fill
                for i in range(4):
                    row[i].border = medium_border
        elif os_name == "Microsoft Windows 10 Pro":
            if win_version == win10_version:
                row[3].fill = green_fill
            else:
                row[3].fill = red_fill
                for i in range(4):
                    row[i].border = medium_border
        else:
            row[3].fill = yellow_fill

    disk_space_rule = DataBarRule(start_type="num",
                                  start_value=1,
                                  end_type="num",
                                  end_value="300",
                                  color="0000FF00")  # Grøn
    worksheet.conditional_formatting.add("E2:E1000", disk_space_rule)

    # Iterate through column E starting from row 2
    for row_num in range(2, worksheet.max_row + 1):
        cell = worksheet.cell(row=row_num, column=5)  # Column E is represented by index 5
        if cell.value is not None and cell.value < 50:
            cell.font = bold_font

    worksheet.freeze_panes = "C2"

    # Gem excel fil
    excel_filename = '0_samlet_data.xlsx'
    excel_path = os.path.join(folder_path, excel_filename)
    workbook.save(excel_path)

    print(f"Data hentet og gemt i: '{excel_filename}'.")

except Exception as e:
    print(f"Der er sket en fejl: {e}")

# Hold cmd-prompt åben
input("Tryk Enter for at afslutte...")
