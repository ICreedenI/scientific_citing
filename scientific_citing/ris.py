import logging
import os
import re

import PyPDF2
import requests
from clipboard import copy
from colorful_terminal import *
from easy_tasks import pickle_unpack, remove_dublicates
from exception_details import print_exception_details


# Set the logging level to ERROR
logging.getLogger("PyPDF2").setLevel(logging.ERROR)


abbreviations = {
    "TY": "Type of reference (must be the first tag; see `abbreviation_to_type_of_reference`)",
    "A1": "Primary Authors (each author on its own line preceded by the A1 tag)",
    "A2": "Secondary Authors (each author on its own line preceded by the A2 tag)",
    "A3": "Tertiary Authors (each author on its own line preceded by the A3 tag)",
    "A4": "Subsidiary Authors (each author on its own line preceded by the A4 tag)",
    "AB": "Abstract",
    "AD": "Author Address",
    "AN": "Accession Number",
    "AU": "Author (each author on its own line preceded by the AU tag)",
    "AV": "Location in Archives",
    "BT": "This field maps to T2 for all reference types except for Whole Book and Unpublished Work references. It can contain alphanumeric characters. There is no practical limit to the length of this field.",
    "C1": "Custom 1",
    "C2": "Custom 2",
    "C3": "Custom 3",
    "C4": "Custom 4",
    "C5": "Custom 5",
    "C6": "Custom 6",
    "C7": "Custom 7",
    "C8": "Custom 8",
    "CA": "Caption",
    "CN": "Call Number",
    "CP": "This field can contain alphanumeric characters. There is no practical limit to the length of this field.",
    "CT": "Title of unpublished reference",
    "CY": "Place Published",
    "DA": "Date",
    "DB": "Name of Database",
    "DO": "DOI",
    "DP": "Database Provider",
    "ED": "Editor",
    "EP": "End Page",
    "ET": "Edition",
    "ID": "Reference ID",
    "IS": "Issue number",
    "J1": "Periodical name: user abbreviation 1. This is an alphanumeric field of up to 255 characters.",
    "J2": "Alternate Title (this field is used for the abbreviated title of a book or journal name, the latter mapped to T2)",
    "JA": "Periodical name: standard abbreviation. This is the periodical in which the article was (or is to be, in the case of in-press references) published. This is an alphanumeric field of up to 255 characters.",
    "JF": "Journal/Periodical name: full format. This is an alphanumeric field of up to 255 characters.",
    "JO": "Journal/Periodical name: full format. This is an alphanumeric field of up to 255 characters.",
    "KW": "Keywords (keywords should be entered each on its own line preceded by the tag)",
    "L1": "Link to PDF. There is no practical limit to the length of this field. URL addresses can be entered individually, one per tag or multiple addresses can be entered on one line using a semi-colon as a separator. These links should end with a file name, and not simply a landing page. Use the UR tag for URL links.",
    "L2": "Link to Full-text. There is no practical limit to the length of this field. URL addresses can be entered individually, one per tag or multiple addresses can be entered on one line using a semi-colon as a separator.",
    "L3": "Related Records. There is no practical limit to the length of this field.",
    "L4": "Image(s). There is no practical limit to the length of this field.",
    "LA": "Language",
    "LB": "Label",
    "LK": "Website Link",
    "M1": "Number",
    "M2": "Miscellaneous 2. This is an alphanumeric field and there is no practical limit to the length of this field.",
    "M3": "Type of Work",
    "N1": "Notes",
    "N2": "Abstract. This is a free text field and can contain alphanumeric characters. There is no practical length limit to this field.",
    "NV": "Number of Volumes",
    "OP": "Original Publication",
    "PB": "Publisher",
    "PP": "Publishing Place",
    "PY": "Publication year (YYYY)",
    "RI": "Reviewed Item",
    "RN": "Research Notes",
    "RP": "Reprint Edition",
    "SE": "Section",
    "SN": "ISBN/ISSN",
    "SP": "Start Page",
    "ST": "Short Title",
    "T1": "Primary Title",
    "T2": "Secondary Title (journal title, if applicable)",
    "T3": "Tertiary Title",
    "TA": "Translated Author",
    "TI": "Title",
    "TT": "Translated Title",
    "U1": "User definable 1. This is an alphanumeric field and there is no practical limit to the length of this field.",
    "U2": "User definable 2. This is an alphanumeric field and there is no practical limit to the length of this field.",
    "U3": "User definable 3. This is an alphanumeric field and there is no practical limit to the length of this field.",
    "U4": "User definable 4. This is an alphanumeric field and there is no practical limit to the length of this field.",
    "U5": "User definable 5. This is an alphanumeric field and there is no practical limit to the length of this field.",
    "UR": "URL",
    "VL": "Volume number",
    "VO": "Published Standard number",
    "Y1": "Primary Date",
    "Y2": "Access Date",
    "ER": "End of Reference (must be empty and the last tag) ",
}
"ris syntax - Long descriptions of abbreviations used in ris-files"

abbreviation_to_type_of_reference = {
    "ABST": "Abstract",
    "ADVS": "Audiovisual material",
    "AGGR": "Aggregated Database",
    "ANCIENT": "Ancient Text",
    "ART": "Art Work",
    "BILL": "Bill",
    "BLOG": "Blog",
    "BOOK": "Whole book",
    "CASE": "Case",
    "CHAP": "Book chapter",
    "CHART": "Chart",
    "CLSWK": "Classical Work",
    "COMP": "Computer program",
    "CONF": "Conference proceeding",
    "CPAPER": "Conference paper",
    "CTLG": "Catalog",
    "DATA": "Data file",
    "DBASE": "Online Database",
    "DICT": "Dictionary",
    "EBOOK": "Electronic Book",
    "ECHAP": "Electronic Book Section",
    "EDBOOK": "Edited Book",
    "EJOUR": "Electronic Article",
    "WEB": "Web Page",
    "ENCYC": "Encyclopedia",
    "EQUA": "Equation",
    "FIGURE": "Figure",
    "GEN": "Generic",
    "GOVDOC": "Government Document",
    "GRANT": "Grant",
    "HEAR": "Hearing",
    "ICOMM": "Internet Communication",
    "INPR": "In Press",
    "JFULL": "Journal (full)",
    "JOUR": "Journal",
    "LEGAL": "Legal Rule or Regulation",
    "MANSCPT": "Manuscript",
    "MAP": "Map",
    "MGZN": "Magazine article",
    "MPCT": "Motion picture",
    "MULTI": "Online Multimedia",
    "MUSIC": "Music score",
    "NEWS": "Newspaper",
    "PAMP": "Pamphlet",
    "PAT": "Patent",
    "PCOMM": "Personal communication",
    "RPRT": "Report",
    "SER": "Serial publication",
    "SLIDE": "Slide",
    "SOUND": "Sound recording",
    "STAND": "Standard",
    "STAT": "Statute",
    "THES": "Thesis/Dissertation",
    "UNBILL": "Unenacted Bill",
    "UNPB": "Unpublished work",
    "VIDEO": "Video recording",
    "WEB": "Web Page ",
}
"ris syntax for the type of the reference"

secondary_abbreviations = {
    "TY": "Type of reference",
    "A1": "Primary authors",
    "A2": "Secondary authors",
    "A3": "Tertiary authors",
    "A4": "Subsidiary authors",
    "AB": "Abstract",
    "AD": "Author address",
    "AN": "Accession number",
    "AU": "Author",
    "AV": "Location in archives",
    "BT": "Secondary title extended",
    "C1": "Custom 1",
    "C2": "Custom 2",
    "C3": "Custom 3",
    "C4": "Custom 4",
    "C5": "Custom 5",
    "C6": "Custom 6",
    "C7": "Custom 7",
    "C8": "Custom 8",
    "CA": "Caption",
    "CN": "Call number",
    "CP": "Alphanumeric characters",
    "CT": "Title of unpublished reference",
    "CY": "Place published",
    "DA": "Date",
    "DB": "Name of database",
    "DO": "DOI",
    "DP": "Database provider",
    "ED": "Editor",
    "EP": "End page",
    "ET": "Edition",
    "ID": "Reference ID",
    "IS": "Issue number",
    "J1": "User abbreviation 1",
    "J2": "Alternate title",
    "JA": "Standard abbreviation",
    "JF": "Journal/Periodical name",
    "JO": "Journal/Periodical name 2",
    "KW": "Keywords",
    "L1": "Link to PDF",
    "L2": "Link to full-text",
    "L3": "Related records",
    "L4": "Image(s)",
    "LA": "Language",
    "LB": "Label",
    "LK": "Website link",
    "M1": "Number",
    "M2": "Miscellaneous 2",
    "M3": "Type of work",
    "N1": "Notes",
    "N2": "Abstract N2",
    "NV": "Number of volumes",
    "OP": "Original publication",
    "PB": "Publisher",
    "PP": "Publishing place",
    "PY": "Publication year",
    "RI": "Reviewed item",
    "RN": "Research notes",
    "RP": "Reprint edition",
    "SE": "Section",
    "SN": "ISBN/ISSN",
    "SP": "Start page",
    "ST": "Short title",
    "T1": "Primary title",
    "T2": "Secondary title",
    "T3": "Tertiary title",
    "TA": "Translated author",
    "TI": "Article title",
    "TT": "Translated title",
    "U1": "User definable 1",
    "U2": "User definable 2",
    "U3": "User definable 3",
    "U4": "User definable 4",
    "U5": "User definable 5",
    "UR": "URL",
    "VL": "Volume number",
    "VO": "Published standard number",
    "Y1": "Primary date",
    "Y2": "Access date",
}
"ris syntax - Short descriptions of abbreviations used in ris-files"

journal_abbreviations_by_CASSI: dict = pickle_unpack(
    os.path.join(
        os.path.dirname(__file__),
        "journal_to_abbreviation_corejournals_by_CASSI.pickle",
    )
)
"Dictionary containing the journal abbreviations by CASSI"


# usage
def first_letter_upper_case(string: str):
    """Formats a string to have every word in lower case except the first letter of each word which is upper case.

    Args:
        string (str): String to be formated

    Returns:
        str: Reformated string.
    """
    words = string.split()  # Split the string into individual words
    formatted_words = []

    for word in words:
        formatted_word = word.capitalize()  # Capitalize the first letter
        formatted_words.append(formatted_word)

    formatted_string = " ".join(formatted_words)  # Join the words back into a string

    return formatted_string


def transform_to_valid_filename(string):
    # Define the invalid characters in Windows filenames
    invalid_chars = r'[<>:"/\\|?*\x00-\x1F]'

    # Replace invalid characters with an underscore
    transformed_string = re.sub(invalid_chars, "_", string)

    # Remove any leading or trailing spaces or periods
    transformed_string = transformed_string.strip(" .")

    # Limit the filename length to 255 characters
    transformed_string = transformed_string[:255]

    return transformed_string


# Function to find and extract DOIs from a PDF
def extract_dois_from_pdf(pdf_path):
    dois = set()  # Use a set to store unique DOIs

    with open(pdf_path, "rb") as pdf_file:
        pdf_reader = PyPDF2.PdfFileReader(pdf_file)

        for page_num in range(pdf_reader.numPages):
            page = pdf_reader.getPage(page_num)
            text = page.extractText()

            # Regular expression to find DOIs (may not capture all DOI formats)
            doi_pattern = r"\b(10\.\d{4,9}/[-._;()/:a-zA-Z0-9]*)\b"

            # Find all matches on the page and add them to the set
            matches = re.findall(doi_pattern, text)
            dois.update(matches)

    return list(dois)


# Function to find and extract the first DOI from a PDF
def extract_first_doi_from_pdf(pdf_path):
    with open(pdf_path, "rb") as pdf_file:
        pdf_reader = PyPDF2.PdfFileReader(pdf_file)

        for page_num in range(pdf_reader.numPages):
            page = pdf_reader.getPage(page_num)
            text = page.extractText()

            # Regular expression to find DOIs (may not capture all DOI formats)
            doi_pattern = r"\b(10\.\d{4,9}/[-._;()/:a-zA-Z0-9]*)\b"

            # Find the first match on the page and return it
            match = re.search(doi_pattern, text)
            if match:
                return match.group()

    return None


# Function to extract the title of a research paper from a PDF
def extract_paper_title(pdf_path):
    with open(pdf_path, "rb") as pdf_file:
        pdf_reader = PyPDF2.PdfFileReader(pdf_file)

        # Iterate through the pages and extract text
        for page_num in range(pdf_reader.numPages):
            page = pdf_reader.getPage(page_num)
            text = page.extractText()

            # Define a regular expression pattern for common title patterns
            title_pattern = (
                r"(?:\bTitle\b\s*:\s*(.*?)(?:\n|$))|" r"(?:^Title\s*\n(.*?)(?:\n|$))"
            )

            match = re.search(title_pattern, text, re.I | re.DOTALL)
            if match:
                title = match.group(1) or match.group(2)
                title = title.strip()
                return title

    return None


def doi_to_ris(
    doi: str,
    filepath: str = None,
    rename_file_to_angewandte_citing_style: bool = False,
    rename_ris_to_title: bool = False,
    move_ris_in_title_dir: bool = False,
    copy_citation: bool = False,
    max_title_length: int = 75,
):
    """Use 'https://api.crossref.org/works/{doi}' to get the regerence information and write the ris file accordingly.

    Args:
        doi (str): doi, the stuff behind the 'doi :'
        filepath (str, optional): Give a filepath to save the ris file, filepath can be the target directory. Defaults to None.
        rename_file_to_angewandte_citing_style (bool, optional): Renames the ris file to the citation as wished by Angewandte Chemie. Defaults to False.
        rename_ris_to_title (bool, optional): Renames the ris file to the title as found for the field T1. Defaults to False.

    Returns:
        str | None: ris file content if proceses was successful otherwise None
    """
    doi = doi.strip()
    doi = doi.replace("https://doi.org/", "")
    doi = doi.replace("DOI: ", "")
    doi = doi.replace("doi: ", "")
    api_url = f"https://api.crossref.org/works/{doi}"

    try:
        response = requests.get(api_url)
        response.raise_for_status()
        data = response.json()

        if data["status"] == "ok":
            work = data["message"]
            title = work.get("title", [])[0]
            authors = work.get("author", [])
            year = work.get("created", {}).get("date-parts", [])[0][0]
            journal = work.get("container-title", [])[0]
            volume = work.get("volume", "")
            issue = work.get("issue", "")
            pages = work.get("page", "")

            ris_data = []

            # Adding required RIS tags
            ris_data.append("TY  - JOUR")
            ris_data.append(f"T1  - {title}")

            # Adding authors
            for author in authors:
                given_name = author.get("given", "")
                surname = author.get("family", "")
                ris_data.append(f"AU  - {surname}, {given_name}")

            # Adding year
            ris_data.append(f"PY  - {year}")

            # Adding journal information
            ris_data.append(f"JO  - {journal}")
            ris_data.append(f"VL  - {volume}")
            ris_data.append(f"IS  - {issue}")
            ris_data.append(f"SP  - {pages}")
            ris_data.append(f"DO  - {doi}")

            ris_data.append("ER  -")

            ris = "\n".join(ris_data)

            ref = RIS(ris)
            name = ref.angewandte_chemie_style(formated=False)
            if copy_citation:
                copy(name.rstrip("."))
            if rename_file_to_angewandte_citing_style:
                if os.path.isdir(filepath):
                    dirpath = filepath
                else:
                    dirpath = os.path.dirname(filepath)
                filepath = os.path.join(dirpath, name + "ris")
            if rename_ris_to_title:
                if os.path.isdir(filepath):
                    dirpath = filepath
                else:
                    dirpath = os.path.dirname(filepath)
                filepath = os.path.join(
                    dirpath,
                    transform_to_valid_filename(ref.entry["T1"][:max_title_length])
                    + ".ris",
                )
            if move_ris_in_title_dir:
                dp = os.path.dirname(filepath)
                filepath = os.path.join(
                    dp,
                    transform_to_valid_filename(ref.entry["T1"][:max_title_length]),
                    os.path.basename(filepath),
                )
                ndp = os.path.dirname(filepath)
                if not os.path.isdir(ndp):
                    os.makedirs(ndp)
            if filepath != None:
                with open(filepath, "w", encoding="utf-8") as f:
                    f.write(ris)

            return ref

    except requests.exceptions.RequestException as e:
        print(f"Error: {e}")

    return None


def pdf_to_ris(
    pdf_path: str,
    ris_dir: str = None,
    rename_ris_to_angewandte_citing_style: bool = False,
    rename_ris_to_title: bool = False,
    rename_pdf_to_angewandte_citing_style: bool = False,
    rename_pdf_to_title: bool = False,
    move_pdf_and_ris_in_title_dir: bool = False,
    copy_citation: bool = False,
    max_title_length: int = 75,
):
    doi = extract_first_doi_from_pdf(pdf_path)
    ris = doi_to_ris(
        doi,
        ris_dir,
        rename_file_to_angewandte_citing_style=rename_ris_to_angewandte_citing_style,
        rename_ris_to_title=rename_ris_to_title,
        move_ris_in_title_dir=move_pdf_and_ris_in_title_dir,
        copy_citation=copy_citation,
        max_title_length=max_title_length,
    )

    if rename_pdf_to_angewandte_citing_style:
        cit = ris.angewandte_chemie_style(formated=False) + ".pdf"
        dp = os.path.dirname(pdf_path)
        nfp = os.path.join(dp, cit)
        os.rename(pdf_path, nfp)
        pdf_path = nfp

    if rename_pdf_to_title:
        title = ris.entry["T1"][:max_title_length]
        title = transform_to_valid_filename(title)
        dp = os.path.dirname(pdf_path)
        nfp = os.path.join(dp, title)
        os.rename(pdf_path, nfp)
        pdf_path = nfp

    if move_pdf_and_ris_in_title_dir:
        title = ris.entry["T1"][:max_title_length]
        dp = os.path.dirname(pdf_path)
        filename = os.path.basename(pdf_path)
        nfp = os.path.join(dp, title, filename)
        os.rename(pdf_path, nfp)
        pdf_path = nfp

    return ris


class RIS:
    """Read RIS files and create dictionaries for the entries of each reference contained in the file.
     Date could be formated year/month/day - maybe that's not a standard
    Instead of a filepath you can input the string content of the RIS file."""

    def __init__(self, filepath: str, encoding="utf-8") -> None:
        self.filecontent = ""
        self.references = []
        "RIS files can contain multiple references which are seperated by 'ER - '. These references will each be an item of the list `references` if form of a dictionary."
        self.entry = {}
        "Dictionary of entries; In case of multiple references it's the last reference only."
        if not os.path.exists(filepath):
            content = filepath
            self.filepath = None
        else:
            self.filepath = filepath
            with open(filepath, "r", encoding=encoding) as f:
                content = f.read()
        self.filecontent = content
        self.type_of_reference = self._classify_reference(content)
        "Fore now classified to be one of: 'Journal Quote', 'Book Quote with Editor', 'Book Quote without Editor', 'Thesis/Dissertation Quote', 'Webpage Quote', 'Other Quote'"
        items = content.split("ER - \n")
        for item in items:
            self.handle_items(item)
        self.cite_as = self.angewandte_chemie_style()

    def handle_items(self, filecontent):
        self.entry = {}
        for line in filecontent.splitlines():
            if line.startswith("#") or "-" not in line:
                continue
            tag = line.split("-")[0].strip()
            entry = line.split("-")[1].strip()
            try:
                e = self.entry[tag.strip()]
                if isinstance(e, list):
                    e.append(entry)
                else:
                    e = [e, entry]
                self.entry[tag] = e
            except:
                self.entry[tag] = entry

        atags = ["A1", "A2", "A3", "A4", "AU"]
        for a in atags:
            try:
                if isinstance(self.entry[a], list) or isinstance(self.entry[a], tuple):
                    self.entry[a] = [first_letter_upper_case(e) for e in self.entry[a]]
                else:
                    self.entry[a] = first_letter_upper_case(self.entry[a])
            except:
                pass

        for abbrev, title in secondary_abbreviations.items():
            try:
                self.entry[title] = self.entry[abbrev]
            except:
                pass
        try:
            self.entry["Type of reference (long)"] = abbreviation_to_type_of_reference[
                self.entry["Type of reference"]
            ]
        except:
            pass
        if self.entry != {}:
            self.references.append(self.entry)
        else:
            self.entry = self.references[0]

    def _classify_reference(self, filecontent: str):
        """Classifies reference to be one of: 'Journal Quote', 'Book Quote with Editor', 'Book Quote without Editor', 'Thesis Quote', 'Dissertation Quote', 'Webpage Quote', 'Other Quote'
        TY  - THES -> Thesis Quote (might be a Dissertation Quote)
        TY  - DISS -> Dissertation Quote
        """
        reference = filecontent
        reference_type = None
        if "TY  - JOUR" in reference:
            reference_type = "Journal Quote"
        elif "TY  - BOOK" in reference:
            if "ED  -" in reference:
                reference_type = "Book Quote with Editor"
            else:
                reference_type = "Book Quote without Editor"
        # elif 'TY  - THES' in reference or 'TY  - DISS' in reference:
        #     reference_type = 'Thesis/Dissertation Quote'
        elif "TY  - THES" in reference:
            reference_type = "Thesis Quote"
        elif "TY  - DISS" in reference:
            reference_type = "Dissertation Quote"
        elif "TY  - PAT" in reference:
            reference_type = "Patent Quote"
        elif "TY  - WEB" in reference:
            reference_type = "Webpage Quote"
        else:
            reference_type = "Other Quote"

        return reference_type

    def _get_authors(self, filecontent: str) -> list:
        authors = []
        record = ""
        for line in filecontent.split("\n"):
            line = line.strip()
            if line.startswith("AU  - "):
                authors.append(line[6:])
            elif line.startswith("A1  - "):
                record += line[6:] + " "
            elif line.startswith("A2  - "):
                record += line[6:] + " "
            elif line.startswith("A3  - "):
                record += line[6:] + " "
            elif line.startswith("A4  - "):
                record += line[6:] + " "
        if record:
            authors.append(record.strip())
        return authors

    def _transform_authors(self, author_list):
        """Get a list containing tuples of the abbreviated first names and the full last name. It is assumed that the names of the authors in the list are written as 'lastname, firstnames'"""
        transformed_authors = []
        for author in author_list:
            names = author.split(", ")
            last_name = names[0].strip()
            first_names = names[1].split()
            abbreviated_first_names = [name[0] + "." for name in first_names[:-1]]
            abbreviated_first_names.append(first_names[-1])
            transformed_authors.append((abbreviated_first_names, last_name))
        return transformed_authors

    def _add_authors_as_abbreviated_first_names_full_last_name(
        self,
        entry: dict,
        string: str,
        max_else_etal: int = 10,
    ):
        """Adds the authors to the string.
         Example: Author is 'Hans Peter Wurst' -> H. P. Wurst

        Args:
            entry (dict): Dictionary (self.entry) containing the informations.
            string (str): string to be added to.
            max_else_etal (int, optional): Maximum amount of authors to list. If amount of authors is greater only the first will be mentioned and other will be replaced by 'et al.'. Defaults to 10.

        Returns:
            str: Edited string
        """
        out = string
        authors = []
        atags = ["A1", "A2", "A3", "A4", "AU"]
        for a in atags:
            if a in entry:
                if isinstance(entry[a], list) or isinstance(entry[a], tuple):
                    authors.extend([e.strip() for e in entry[a]])
                else:
                    authors.append(entry[a].strip())

        authors = remove_dublicates([a.strip() for a in authors if a.strip()])

        if len(authors) < max_else_etal:
            for a in authors:
                if "," in a:
                    lastname, firstnames = a.split(", ")
                    firstnames = firstnames.split(" ")
                    out += (
                        " ".join([c[0] + "." for c in firstnames])
                        + " "
                        + lastname
                        + ", "
                    )
                else:
                    out += a + ", "
        else:
            if "," in authors[0]:
                lastname, firstnames = authors[0].split(", ")
                firstnames = firstnames.split(" ")
                out += (
                    " ".join([c[0] + "." for c in firstnames])
                    + " "
                    + lastname
                    + " et al., "
                )
            else:
                out += authors[0] + " et al., "
        return out

    def _add_authors_as_full_last_name_abbreviated_first_names(
        self,
        entry: dict,
        string: str,
        max_else_etal: int = 10,
    ):
        """Adds the authors to the string.
         Example: Author is 'Hans Peter Wurst' -> Wurst, H. P.

        Args:
            entry (dict): Dictionary (self.entry) containing the informations.
            string (str): string to be added to.
            max_else_etal (int, optional): Maximum amount of authors to list. If amount of authors is greater only the first will be mentioned and other will be replaced by 'et al.'. Defaults to 10.

        Returns:
            str: Edited string
        """
        out = string
        authors = []
        atags = ["A1", "A2", "A3", "A4", "AU"]
        for a in atags:
            if a in entry:
                if isinstance(entry[a], list) or isinstance(entry[a], tuple):
                    authors.extend([e.strip() for e in entry[a]])
                else:
                    authors.append(entry[a].strip())

        authors = remove_dublicates([a.strip() for a in authors if a.strip()])

        if len(authors) < max_else_etal:
            for a in authors:
                if "," in a:
                    lastname, firstnames = a.split(", ")
                    firstnames = firstnames.split(" ")
                    out += (
                        lastname
                        + " "
                        + " ".join([c[0] + "." for c in firstnames])
                        + ", "
                    )
                else:
                    out += a + ", "
        else:
            if "," in authors[0]:
                lastname, firstnames = authors[0].split(", ")
                firstnames = firstnames.split(" ")
                out += (
                    lastname
                    + ", "
                    + " ".join([c[0] + "." for c in firstnames])
                    + " et al., "
                )
            else:
                out += authors[0] + " et al., "
        return out

    def _get_editors_as_abbreviated_first_names_full_last_name(
        self,
        entry: dict,
    ):
        eds = []
        if isinstance(entry["ED"], str):
            eds.append(
                entry["ED"].split(", ")[-1][0] + ". " + entry["ED"].split(", ")[0]
            )
        else:
            for e in entry["ED"]:
                eds.append(e.split(", ")[-1][0] + ". " + e.split(", ")[0])
        return ", ".join(eds)

    def _add_year(
        self, entry: dict, string: str, bold: bool = True, italic: bool = False
    ):
        out = string
        if bold:
            format_in = Style.BOLD
            format_out = Style.NOT_BOLD
        elif italic:
            format_in = Style.ITALIC
            format_out = Style.NOT_ITALIC
        else:
            format_in = format_out = ""

        out += format_in + entry.get("Publication year", "") + format_out + ", "

        return out

    def _add_volume_number(
        self, entry: dict, string: str, bold: bool = True, italic: bool = False
    ):
        out = string
        if bold:
            format_in = Style.BOLD
            format_out = Style.NOT_BOLD
        elif italic:
            format_in = Style.ITALIC
            format_out = Style.NOT_ITALIC
        else:
            format_in = format_out = ""

        out += format_in + entry.get("Volume number", "") + format_out + ", "

        return out

    def _add_publisher(
        self, entry: dict, string: str, bold: bool = True, italic: bool = False
    ):
        out = string
        if bold:
            format_in = Style.BOLD
            format_out = Style.NOT_BOLD
        elif italic:
            format_in = Style.ITALIC
            format_out = Style.NOT_ITALIC
        else:
            format_in = format_out = ""

        out += format_in + entry.get("Publisher", "") + format_out + ", "

        return out

    def _add_publishing_place(
        self, entry: dict, string: str, bold: bool = True, italic: bool = False
    ):
        out = string
        if bold:
            format_in = Style.BOLD
            format_out = Style.NOT_BOLD
        elif italic:
            format_in = Style.ITALIC
            format_out = Style.NOT_ITALIC
        else:
            format_in = format_out = ""

        out += format_in + entry.get("Publishing Place", "") + format_out + ", "

        return out

    def _add_title(
        self, entry: dict, string: str, bold: bool = True, italic: bool = False
    ):
        out = string
        if bold:
            format_in = Style.BOLD
            format_out = Style.NOT_BOLD
        elif italic:
            format_in = Style.ITALIC
            format_out = Style.NOT_ITALIC
        else:
            format_in = format_out = ""

        try:
            title = (
                entry.get("TI")
                or entry.get("T1")
                or entry.get("T2")
                or entry.get("T3")
                or entry.get("ST")
            )
        except KeyError:
            raise Exception("No title found")

        out += format_in + title + format_out + ", "

        return out

    def _add_type_of_reference(
        self, entry: dict, string: str, bold: bool = True, italic: bool = False
    ):
        out = string
        if bold:
            format_in = Style.BOLD
            format_out = Style.NOT_BOLD
        elif italic:
            format_in = Style.ITALIC
            format_out = Style.NOT_ITALIC
        else:
            format_in = format_out = ""

        type_of_reference = abbreviation_to_type_of_reference.get(
            entry.get("Type of reference", "")
        )

        out += format_in + type_of_reference + format_out + ", "

        return out

    def _add_pages(
        self,
        entry: dict,
        string: str,
        include_end_page: bool = False,
        pre_page_string: str = "",
        bold: bool = True,
        italic: bool = False,
    ):
        out = string
        if bold:
            format_in = Style.BOLD
            format_out = Style.NOT_BOLD
        elif italic:
            format_in = Style.ITALIC
            format_out = Style.NOT_ITALIC
        else:
            format_in = format_out = ""

        start_page = entry.get("Start Page", "")
        end_page = entry.get("End Page", "")
        if not include_end_page:
            out += format_in + pre_page_string + start_page + format_out
        elif start_page:
            out += format_in + pre_page_string + start_page + end_page + format_out

        return out

    def _add_journal_or_periodical_name(
        self,
        entry: dict,
        string: str,
        characters_to_delete: str = ",",
        bold: bool = True,
        italic: bool = False,
    ):
        out = string
        if bold:
            format_in = Style.BOLD
            format_out = Style.NOT_BOLD
        elif italic:
            format_in = Style.ITALIC
            format_out = Style.NOT_ITALIC
        else:
            format_in = format_out = ""

        jrnl = entry.get("Journal/Periodical name") or entry.get(
            "Journal/Periodical name 2"
        )
        jrnl: str
        for char in characters_to_delete:
            jrnl = jrnl.replace(char, "")
        out += format_in + jrnl + format_out + ", "

        return out

    def angewandte_chemie_style(
        self, entry=None, formated: bool = True, **kwargs
    ) -> str:
        """Generate the citation in the style of the Angewandte Chemie.

        Args:
            entry (_type_, optional): _description_. Defaults to None.
            formated (bool, optional): _description_. Defaults to True.
            **kwargs:
                - Use dissertation = True to indicate that it is a dissertation if the reference type is 'Thesis/Dissertation'.
                - Use master = True to indicate that it is a master thesis if the reference type is 'Thesis/Dissertation'.

        Raises:
            Exception: _description_

        Returns:
            str: _description_
        """
        if entry is None:
            entry = self.entry

        kwargs.setdefault("dissertation", False)
        kwargs.setdefault("master", False)

        if formated:
            bold = Style.BOLD
            not_bold = Style.NOT_BOLD
            italic = Style.ITALIC
            not_italic = Style.NOT_ITALIC
        else:
            bold = not_bold = italic = not_italic = ""

        out = ""

        out = self._add_authors_as_abbreviated_first_names_full_last_name(
            entry, out, 10
        )

        jrnl = (
            entry.get("Journal/Periodical name")
            or entry.get("Journal/Periodical name 2")
            or entry.get("Secondary title extended")
        )
        assert isinstance(jrnl, str)  # no journal or book name

        type_of_reference = abbreviation_to_type_of_reference.get(
            entry.get("Type of reference", "")
        )

        if type_of_reference == "Thesis/Dissertation":
            if kwargs["dissertation"]:
                out += f"Dissertation, {jrnl}, "
            elif kwargs["master"]:
                out += f"Master, {jrnl}, "
            else:
                out += f"Thesis/Dissertation, {jrnl}, "
            out += bold + entry.get("Publication year", "") + not_bold + ", "

        elif type_of_reference in (
            "Whole book",
            "Book chapter",
            "Electronic Book",
            "Electronic Book Section",
            "Edited Book",
        ):
            try:
                entry["ED"]
                out = italic + out.rstrip(", ") + f" in {jrnl}" + not_italic + ", "
                if entry.get("Volume number", "") != "":
                    out += (
                        italic
                        + "Vol. "
                        + entry.get("Volume number", "")
                        + not_italic
                        + ", "
                    )
                out += (
                    "(Editor: "
                    + self._get_editors_as_abbreviated_first_names_full_last_name(entry)
                    + "), "
                )
                out += entry.get("Publisher", "") + ", "
                out += entry.get("Publishing Place", "") + ", "
                out += bold + entry.get("Publication year", "") + not_bold + ", "
            except Exception as e:
                print_exception_details(e)
                out += italic + jrnl.replace(",", "") + not_italic + ", "
                if entry.get("Volume number", "") != "":
                    out += "Vol. " + entry.get("Volume number", "") + ", "
                out += entry.get("Publisher", "") + ", "
                out += entry.get("Publishing Place", "") + ", "
                out += bold + entry.get("Publication year", "") + not_bold + ", "

        else:
            out += italic + jrnl.replace(",", "") + not_italic + ", "
            out += bold + entry.get("Publication year", "") + not_bold + ", "
            out += italic + entry.get("Volume number", "") + not_italic + ", "

        out = out.replace(", ,", ",")

        start_page = entry.get("Start Page", "")
        end_page = entry.get("End Page", "")
        if (
            type_of_reference == "Journal"
            or type_of_reference == "Journal (full)"
            or type_of_reference == "Electronic Article"
        ):
            out += start_page
        elif start_page:
            out += "S. " + start_page + "–" + end_page

        out = out.rstrip(", ") + "."

        return out

    def cite_by_rules(self, rules: dict[str, dict[str, dict[str, str]]]):
        pass

    def rename_pdf_to_angewandte_citing_style(self):
        if self.filepath:
            new_filename = self.angewandte_chemie_style(formated=False) + "pdf"
            new_filename = transform_to_valid_filename(new_filename)
            new_filepath = os.path.join(os.path.dirname(self.filepath), new_filename)
            os.rename(self.filepath, new_filepath)
            self.filepath = new_filepath
            return new_filepath


def chain_references_to_string_list_qutotes_Angewandte_Chemie(references: list[RIS]):
    refs = []
    for r in references:
        for rr in r.references:
            re = r.angewandte_chemie_style(rr)
            refs.append(re)
    out = ""
    for i, r in enumerate(refs):
        out += f"[{i+1}]\t{r}\n"
    return out.strip()
