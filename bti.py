""" bti.py: A script for retrieving documents from the USPTO Bulk Data Storage System. """

import argparse, datetime, json, multiprocessing, os, queue, re, ssl, socket, sqlite3
import sys, time, urllib.request, webbrowser
os.environ['TK_SILENCE_DEPRECATION'] = '1'
from tkinter import ttk, Tk, Tcl, Label, Entry, Button, Checkbutton, Listbox, BooleanVar
from tkinter import W, E, SINGLE, NONE, NORMAL, DISABLED, CENTER, END
from tkinter import filedialog

class BDSS_Interface:
  """ An interface for the USPTO Bulk Data Storage System. """

  # top-level functions

  @classmethod
  def run_from_command_line(cls):
    """ Processes command-line arguments and calls appropriate functions. """

    # parse command-line arguments
    parser = argparse.ArgumentParser(description='Retrieves PDFs from the USPTO BDSS repository using the BDSS TAR Index.')
    parser.add_argument('ref', type=str, nargs='?', help='Reference number of patent, publication, or patent application.')
    parser.add_argument('-t', type=str, choices=['patent', 'publication', 'application'], help='Type of reference number ("patent", "publication", or "application").')
    parser.add_argument('-d', action='store_true', help='Only downloads document; does not open document in default PDF viewer.')
    parser.add_argument('-q', action='store_true', help='Quiet (suppresses console output).')
    args = parser.parse_args()
    ref = vars(args).get('ref', None)
    document_type = vars(args).get('t', None)
    download_only = (vars(args).get('d', False) is True)
    output = (vars(args).get('q', False) is False)

    if ref is None:           # run GUI
      # check for outdated version
      tk = Tk()
      try:
        tk_version_string = str(Tcl().call("info", "patchlevel"))
        tk_version_match = float(re.search(r'(\d+(\.\d+)?)\.\d+', tk_version_string).group(1))
        if tk_version_match < 8.6:
          print('\nWarning: Your version of Python is running an outdated version of Tcl/Tk.')
          print('As a result, BTI cannot run in GUI mode. BTI can still run as a terminal')
          print('or command-line application to retrieve and show documents.\n')
          print('For more information: https://www.python.org/download/mac/tcltk/\n')
          sys.exit(1)
      except:
        pass

      BTIWindow(tk)
      tk.lift()
      tk.mainloop()
    else:                     # run command-line process
      cls.print_conditional(output, 'BDSS_Interface written by David Stein (mail@usptodata.com)')
      result, documents_path = cls.determine_documents_path(output = output)
      if result is False:
        return
      result = cls.fetch(documents_path, ref, document_type, output = output)
      cls.print_conditional(output, ('Error' if result[0] is False else 'Success') + f': {result[1]}')
      if result[0] is True and download_only is False:
        cls.open_document(result[2])

  @classmethod
  def fetch(cls, documents_path, ref, document_type = None, output = True):
    """ Fetches a document from BDSS and saves it to a local documents path.

        Args:

          documents_path (string): The location to store documents.

          ref (string): An identifier of a reference. This can be a U.S. patent number,
                a U.S. publication number, or a U.S. application number.

                No particular formatting is required; this function regularizes the
                format of the identifier.

                Prefixes: This function (and BTI generally) supports many document
                types: utility, design, plant, reissue, statutory invention
                registration, defensive publication, "additional improvements,"
                X-patents, and reissue X-patents. The reference can include the prefix
                for the type (D, PP, RE, H, T, AI, X, or RX).

                Kind codes: The identifiers for patents and publications should include
                a kind code (e.g., "A1" for first pre-grant publications, and "B1" or
                "B2" for patents). If not included, a kind code will be guessed and
                appended.

          document_type (string, optional): An indicator of the document type - can be
                "patent," "publication," or "application." If omitted, the document type
                is inferred from the format of the reference identifier. Typically, this
                works fine - e.g., if the identifier includes 11 digits, then it must be
                a publication (like "2016/0123456"). But some identifiers could be either
                a patent number or an application number (e.g.: "10123456" could be
                either U.S. Patent No. 10,123,456 or U.S. App. No. 10/123,456). This
                parameter allows the document type to be indicated rather than inferred.

          output (bool, optional): An indicator of whether to provide or suppress
                console output.

        Returns:

          param1 (bool): An indicator of whether the request succeeded or failed.

          param2 (string):
            If param1 is True: A message indicating whether the file already existed in
                  the documents_path folder or was retrieved from BDSS.
            If param1 is False: An error message.

          param3 (string):
            If param1 is True: The complete filename of the stored document.
            If param1 is False: None.
    """

    # process input
    result, reduced_ref, document_type = cls.determine_reference_type(ref, document_type)
    if result is None:
      return (False, f'Unable to determine file type of {ref}.', None)

    # retrieve database entry
    database_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'bti.sqlite')
    if os.path.isfile(database_file):
      cls.print_conditional(output, 'Looking up BDSS information in bti.sqlite')
      result = cls.retrieve_ref_from_database(reduced_ref, document_type)
    else:
      cls.print_conditional(output, 'Querying BTI API for BDSS information')
      result = cls.retrieve_ref_from_bti_api(reduced_ref, document_type)
    if result[0] is False:
      return result
    reduced_ref, document_type, url, offset, size = result[1:]
    if reduced_ref is None or document_type is None:
      return (False, f'No index entry for {ref}.', None)
    cls.print_conditional(output, f'PDF: URL {url}, offset {offset}, size {size}')

    # determine filename and check if file is cached
    local_filename = cls.determine_local_filename(reduced_ref, document_type)
    if local_filename is False:
      return (False, f'Could not determine filename for {ref}.', None)
    local_filename = os.path.join(documents_path, local_filename)
    if os.path.isfile(local_filename) is True:
      return (True, f'{document_type.title()} {ref} already exists.', local_filename)

    # retrieve document
    cls.print_conditional(output, 'Retrieving PDF from BDSS')
    result = cls.retrieve_document_from_bdss(url, offset, size, local_filename)
    return (result[0], result[1], local_filename)

  # mid-level task functions

  @classmethod
  def determine_reference_type(cls, ref, document_type = None):
    """ Determines the type of a reference identifier and returns a regularized
        version of the identifier by which the document could be found in the BTIWindow
        database.

        Args:
          ref (string): The reference identifier.
          document_type (string, optional): An indicator of the document type - can be
                "patent," "publication," or "application."

        Returns:

          param1 (bool): An indicator of whether the determination succeeded or failed.

          param2 (string):
            If param1 is True: The complete filename of the stored document.
            If param1 is False: None.

          param3 (string):
            If param1 is True: An indicator of the document type - either "patent,"
                "publication," or "application."
            If param1 is False: None.
    """

    # remove leading 'US' and patent kind codes, but keep publication kind codes
    reduced_ref = re.sub(r'U|S|B\d+|\.|/|,', '', re.sub(r'\s', '', ref.upper()))
    if document_type is None:  # infer file type based on format
      if len(reduced_ref) >= 11:  # publication (e.g., 20060123456)
        document_type = 'publication'
      # Inference logic: Presume that ref is a patent number if the first two digits are
      # below 12, and an application if above.
      elif ref.find('/') > -1 or (len(reduced_ref) == 8 and all(reduced_ref.startswith(c) is False for c in ['P', 'D', 'T', 'A', 'R', 'X']) and int(reduced_ref[:2]) >= 12):
        document_type = 'application'
      else:
        document_type = 'patent'
    if document_type == 'patent':
      reduced_ref = cls.format_patent_number(reduced_ref, human_readable = False)
    elif document_type == 'publication':
      reduced_ref = cls.format_publication_number(reduced_ref, human_readable = False)
    elif document_type == 'application':
      reduced_ref = cls.format_application_number(reduced_ref, human_readable = False)
    return (False, None, None) if reduced_ref is None else (True, reduced_ref, document_type)

  @classmethod
  def retrieve_ref_from_database(cls, ref, document_type):
    """ Determines the location of a reference within BDSS using a local BTI database.

        Args:
          ref (string): The reference identifier.
          document_type (string, optional): An indicator of the document type - can be
                "patent," "publication," or "application."

        Returns:

          param1 (bool): An indicator of whether the request succeeded or failed.

          param2 (string):
            If param1 is True: A regularized version of the identifier.
            If param1 is False: An error message.

          param3 (string):
            If param1 is True: An indicator of the document type - can be
                "patent," "publication," or "application."
            If param1 is False: None.

          param4 (string):
            If param1 is True: A URL of the BTI .tar file containing the document.
            If param1 is False: None.

          param5 (int):
            If param1 is True: The start location of the document in the .tar file.
            If param1 is False: None.

          param6 (int):
            If param1 is True: The size of the document in the .tar file.
            If param1 is False: None.

    """

    reduced_ref = ref.replace(',', '').replace('/', '')
    year_part = date_part = offset = size = None
    database_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'bti.sqlite')
    conn = sqlite3.connect(database_file)
    result = None
    if document_type in ('application', 'publication'):
      number_type = 'application_number' if document_type == 'application' else 'number'
      query = f'select number, year_part, date_part, offset, size from publications, publications_fts5, date_codes where {number_type} match ? and publication = id and date_code = date_codes.code limit 1'
      result = conn.execute(query, (reduced_ref,)).fetchone()
      if result is not None:
        document_type = 'publication'
        reduced_ref, year_part, date_part, offset, size = result
    if document_type in ('application', 'patent'):
      number_type = 'application_number' if document_type == 'application' else 'number'
      query = f'select number, year_part, date_part, offset, size from patents, patents_fts5, date_codes where {number_type} match ? and patent = id and date_code = date_codes.code limit 1'
      result = conn.execute(query, (reduced_ref,)).fetchone()
      if result is not None:
        document_type = 'patent'
        reduced_ref, year_part, date_part, offset, size = result
    conn.close()
    if result is None:
      return (False, f'No data found for {ref}.', None, None, None, None)
    urls = {
      'patent': 'https://bulkdata.uspto.gov/data/patent/grant/multipagepdf/YEAR/grant_pdf_DATE.tar',
      'publication': 'https://bulkdata.uspto.gov/data/patent/application/multipagepdf/YEAR/app_pdf_DATE.tar'
    }
    url = urls[document_type].replace('YEAR', year_part).replace('DATE', date_part)
    return (True, reduced_ref, document_type, url, offset, size)

  @staticmethod
  def retrieve_ref_from_bti_api(ref, document_type):
    """ Determines the location of a reference within BDSS using the BTI API.

        This function accepts the same arguments and returns the same values as
        retrieve_ref_from_database().
    """

    try:
      socket.setdefaulttimeout(2)
      ssl._create_default_https_context = ssl._create_unverified_context
      timestamp = datetime.datetime.now().strftime('%Y%m%d%H%M%S%')
      url = f'https://www.usptodata.com/bti/bti.php?ref={ref}&type={document_type}&time={timestamp}'
      user_agent = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_3) AppleWebKit/537.75.14 (KHTML, like Gecko) Version/7.0.3 Safari/7046A194A'
      request = urllib.request.Request(url, method='GET', headers={'User-Agent': user_agent})
      with urllib.request.urlopen(request) as request:
        response = json.loads(request.read())
        reduced_ref = response.get('reduced_ref', None)
        document_type = response.get('type', None)
        url = response.get('url', None)
        offset = response.get('offset', None)
        size = response.get('size', None)
        return (True, reduced_ref, document_type, url, offset, size)
    except Exception as e:
      return (False, f'Error while querying BTI API: {e}', None, None, None, None)

  @staticmethod
  def retrieve_document_from_bdss(url, offset, size, local_filename):
    """ Retrieves a document from a BDSS .tar file and locally saves the document.

        Args:
          url (string): The URL of the BTI file.
          offset (int): The offset of the document within the BTI file.
          size (int): The size of the document.
          local_filename (string): The complete filename of the document to be saved.

        Returns:

          param1 (bool): An indicator of whether the determination succeeded or failed.

          param2 (string):
            If param1 is True: A message indicating that the file was successfully
              retrieved.
            If param1 is False: An error message.
    """

    try:
      socket.setdefaulttimeout(2)
      user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2486.0 Safari/537.36 Edge/13.10586'
      headers = {
        'User-Agent': user_agent,
        'Content-Type': 'application/octet-stream',
        'Range': f'bytes={offset}-{offset + size}',
        'Connection': 'close'
      }
      chunk_size = 65536; num_chunks = 0
      request = urllib.request.Request(url, headers=headers)
      with urllib.request.urlopen(request) as response:
        with open(local_filename, 'wb') as f:
          while True:
            data = response.read(chunk_size)
            num_chunks += 1
            if len(data) > 0:
              f.write(data)
            if len(data) < chunk_size or num_chunks * chunk_size > size:
              break
        return (True, f'Retrieved {os.path.basename(local_filename)} from BDSS.', local_filename)
    except Exception as e:
      return (False, f'Exception while retrieving PDF from BDSS: {e}')

  # low-level task functions

  @classmethod
  def determine_documents_path(cls, output = True):
    """ Determines the local document path, either based on an entry in a config
        file or a default path.

        Args:
          output (bool, optional): An indicator of whether to provide or suppress
                console output.

        Returns:

          param1 (bool): An indicator of whether the determination succeeded or failed.

          param2 (string):
            If param1 is True: A message indicating the documents path.
            If param1 is False: An error message.
    """

    # verify existence of documents path
    if os.path.isfile('bti_config.txt'):
      with open('bti_config.txt', 'rt', encoding='UTF-8') as f:
        ini_file = json.loads(f.readline())
        documents_path = ini_file.get('documents_path', None)
        if documents_path is not None:
          if os.path.isdir(documents_path) is True:
            return (True, documents_path)
          cls.print_conditional(output, f'Warning: Documents path {documents_path} does not exist. Using default documents path instead.')

    # create default path if it does not exist
    documents_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'documents')
    if os.path.isdir(documents_path) is False:
      os.mkdir(documents_path)
      if os.path.isdir(documents_path) is False:
        return (False, 'Documents subfolder does not exist and could not be created.')
    cls.print_conditional(output, f'Using default documents path: {documents_path}.')

    return (True, documents_path)

  @classmethod
  def determine_local_filename(cls, ref, document_type):
    """ Determines a standardized local filename of a document.

        Args:
          ref (string): An identifier of a reference. This can be a U.S. patent number,
                a U.S. publication number, or a U.S. application number.
          document_type (string, optional): An indicator of the document type - can be
                "patent," "publication," or "application."

        Returns:
          If the request succeeded, returns the local filename.
          If the request failed, returns None.
    """

    if document_type == 'patent':
      reduced_ref = cls.format_patent_number(ref, human_readable = True)
      if reduced_ref is not None:
        return f'U.S. Patent No. {reduced_ref}.pdf'
    elif document_type == 'publication':
      reduced_ref = cls.format_publication_number(ref, human_readable = False)
      if reduced_ref is not None:
        return f'U.S. Pub. No. {reduced_ref}.pdf'
    elif document_type == 'Application':
      reduced_ref = cls.format_application_number(ref, human_readable = False)
      if reduced_ref is not None:
        return f'U.S. App. No. {reduced_ref}.pdf'
    return None

  @staticmethod
  def check_bdss_online_status():
    """ Determines whether BDSS is online.

        Returns:

          param1 (bool): An indicator of whether BDSS is online.

          param2 (string):
            If param1 is True: An 'Online' message.
            If param1 is False: An error message.
    """

    socket.setdefaulttimeout(2)
    response = None
    try:
      user_agent = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_3) AppleWebKit/537.75.14 (KHTML, like Gecko) Version/7.0.3 Safari/7046A194A'
      request = urllib.request.Request('https://bulkdata.uspto.gov', method='HEAD', headers={'User-Agent': user_agent})
      with urllib.request.urlopen(request) as response:
        return (True, 'Online') if response.status == 200 else (False, response.status)
    except Exception as e:
      return (False, response.status if response is not None else f'Error: {e}')

  @staticmethod
  def check_script_status(url, local_file):
    """ Determines whether the local bti.py script is the latest version.

        Args:
          local_file: The path of the local script file.

        Returns:

          param1 (bool): An indicator of whether the local script is the latest version.

          param2 (string):
            If param1 is True: A message indicating the result of the determination.
            If param1 is False: An error message.
    """

    try:

      # verify that local file exists
      if os.path.isfile(local_file) is False:
        return (True, f'No local file: {local_file}')

      # get remote script date/time
      socket.setdefaulttimeout(2)
      ssl._create_default_https_context = ssl._create_unverified_context
      user_agent = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_3) AppleWebKit/537.75.14 (KHTML, like Gecko) Version/7.0.3 Safari/7046A194A'
      request = urllib.request.Request(url, method='HEAD', headers={'User-Agent': user_agent})
      with urllib.request.urlopen(request) as response:
        remote_script_date_string = response.headers.get('Last-Modified')
      remote_script_date = datetime.datetime.strptime(remote_script_date_string, '%a, %d %b %Y %H:%M:%S %Z').replace(microsecond=0)

      # get local script date/time and compare
      local_script_date = os.path.getmtime(local_file)
      local_script_date = datetime.datetime.fromtimestamp(local_script_date).replace(microsecond=0)
      return (True, 'OK' if local_script_date >= remote_script_date else 'New Version')

    except Exception as e:
      return (False, f'Error: {e}')

  @staticmethod
  def check_database_status(local_database):
    """ Determines whether a local bti.sqlite database exists.

        Args:
          local_database: The path of the local database.

        Returns:

          param1 (bool): An indicator of whether the database is stored locally.

          param2 (string):
            If param1 is True: A message indicating the result of the determination.
            If param1 is False: An error message.
    """

    try:
      response = 'Using BDSS TAR Index API'
      if os.path.isfile(local_database) is True:
        response = 'Using local bti.sqlite database'
      return (True, response)
    except Exception as e:
      return (False, f'Error: {e}')

  @staticmethod
  def fetch_file(url, local_file, output_queue, command):
    """ Retrieves a file at a URL and locally stores it to a specified location.

        Args:
          url (string): The URL of the file.
          local_file (string): The complete path where the file is to be stored.
          output_queue (multiprocessing.Queue): A queue to receive status messages.
          command (string): A command string to include in the output queue.

        Returns:

          param1 (bool): An indicator of whether the file was successfully retrieved.

          param2 (string):
            If param1 is True: An 'OK' message.
            If param1 is False: An error message.
    """
    chunk_size = 65536 * 16; read_size = 0; last_message = time.time()
    temp_file = f'{local_file}.tmp'
    try:
      if os.path.isfile(temp_file):
        os.unlink(temp_file)
      socket.setdefaulttimeout(2)
      ssl._create_default_https_context = ssl._create_unverified_context
      with urllib.request.urlopen(url) as request:
        with open(temp_file, 'wb') as f:
          while True:
            chunk = request.read(chunk_size)
            f.write(chunk)
            if len(chunk) < chunk_size:
              break
            read_size += len(chunk)
            if time.time() - last_message > 0.5:
              last_message = time.time()
              output_queue.put((command, f'Retrieving: {int(read_size / 1024):,} kb', False))
      if os.path.isfile(local_file):
        os.unlink(local_file)
      os.rename(temp_file, local_file)
      # set local file time to same as remote file time
      remote_file_date_string = request.headers.get('Last-Modified')
      remote_file_date = datetime.datetime.strptime(remote_file_date_string, '%a, %d %b %Y %H:%M:%S %Z').replace(microsecond=0)
      os.utime(local_file, (remote_file_date.timestamp(), remote_file_date.timestamp()))
      return (True, 'OK')
    except Exception as e:
      if os.path.isfile(temp_file):
        os.unlink(temp_file)
      return (False, f'Error: {e}')

  # low-level utility functions

  @staticmethod
  def open_document(filename):
    show_pdf_command = 'start ""' if os.name == 'nt' else 'open'
    os.system(f'{show_pdf_command} \"{filename}\"')

  @staticmethod
  def format_application_number(number, human_readable):
    if number is None:
      return None
    number = re.sub(r'U|S|\.|/|,', '', re.sub(r'\s', '', number.upper())).strip()
    if number[0] == 'D':
      application_number = number[1:].lstrip('0').rjust(6, '0')
      number = f'29/{application_number[0:3]},{application_number[4:6]}'
    elif len(number) == 8:
      number = f'{number[:2]}/{number[2:5]},{number[5:]}'
    if human_readable is False:
      number = number.replace('/', '').replace(',', '')
    return number

  @staticmethod
  def format_patent_number(number, human_readable):
    if number is None:
      return None
    number = re.sub(r'U|S|B\d+|\.|/|,', '', re.sub(r'\s', '', number.upper())).strip()
    pn_formatted = re.search(r'(D|PP|RE|H|T|X|RX|AI)?(\d+)', number)
    if pn_formatted is None:
      return None
    prefix = pn_formatted.group(1) or ''
    numeric_part = int(pn_formatted.group(2))
    return f'{prefix}{numeric_part:,}' if human_readable is True else f'{prefix}{numeric_part}'

  @staticmethod
  def format_publication_number(number, human_readable):
    if number is None:
      return None
    number = re.sub(r'U|S|\.|/|,', '', re.sub(r'\s', '', number.upper())).strip()
    if len(number) < 11:
      return None
    if len(number) == 11:     # no series code - append default series code A1
      number = f'{number}A1'
    return number if human_readable is False else f'{number[:4]}/{number[4:11]} {number[11:]}'

  @staticmethod
  def print_conditional(condition, s):
    if condition:
      print(s)

class WorkerProcess:
  """ A worker process class for the BTI Window. """

  @staticmethod
  def run(command, command_arguments, output_queue):
    """ Processes a command and returns results in output_queue. """

    remote_script = 'https://www.usptodata.com/bti/bti.py'

    # monitor command queue
    if command == 'status':
      _, bdss_online_status = BDSS_Interface.check_bdss_online_status()
      output_queue.put(('bdss_status', bdss_online_status, False))
      local_script = os.path.realpath(__file__)
      _, script_status = BDSS_Interface.check_script_status(remote_script, local_script)
      output_queue.put(('script_status', script_status, False))
      local_database = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'bti.sqlite')
      _, local_database_status = BDSS_Interface.check_database_status(local_database)
      output_queue.put(('db_status', local_database_status, True))
      output_queue.put(('status', BDSS_Interface.check_database_status(local_database) + '.', False))
      output_queue.put(('status', 'BDSS is online.' if bdss_online_status == 'Online' else 'Warning: BDSS is offline.', True))
    elif command == 'fetch_script':
      output_queue.put((command, 'Retrieving script', False))
      result = BDSS_Interface.fetch_file(remote_script, os.path.realpath(__file__), output_queue, 'fetch_script')
      output_queue.put((command, result[1], True))
    elif command == 'fetch':
      ref, documents_path, open_document = command_arguments
      result = BDSS_Interface.fetch(documents_path, ref, document_type = None, output = False)
      output_queue.put((command, result[1], True))
      if result[0] is True and open_document is True:
        BDSS_Interface.open_document(result[2])

class BTIWindow:
  """ A Tkinter-based GUI for BTI. """

  # main window functions

  def __init__(self, tk):
    result, documents_path = BDSS_Interface.determine_documents_path(output = True)
    self.documents_path = documents_path if result is True else ''
    self.tk = tk
    self.create_window()
    self.output_queue = multiprocessing.Queue()
    self.tk.after(100, self.check_output_queue)
    self.workers = {}
    self.start_worker('status')

  def create_window(self):
    """ Generates a Tkinter window interface for BTI. """

    # create main window
    self.tk.title('BTI - BDSS TAR Index')
    width = 760; height = 230; font_1 = ('Arial', 12); font_2 = ('Arial', 10)
    self.tk.minsize(width, height); self.tk.maxsize(width, height)

    # create tabs
    tab_control = ttk.Notebook(self.tk)
    tab_main = ttk.Frame(tab_control)
    tab_control.add(tab_main, text='Main')
    tab_configure = ttk.Frame(tab_control)
    tab_control.add(tab_configure, text='Configure')
    tab_control.pack(expand=1, fill="both")

    # main tab
    Label(tab_main, borderwidth=12, text='').grid(row=1, column=0, sticky=W)
    label_text = 'Enter a patent number (x,xxx,xxx), publication number (xxxx/xxxxxxx), or application number (xx/xxxxxx).'
    Label(tab_main, text=label_text, font=font_1).grid(row=1, column=1, columnspan=3, sticky=W)
    self.entry = Entry(tab_main, font=font_1, borderwidth=3)
    self.entry.grid(row=2, column=1, sticky=W+E)
    self.bool_open_document = BooleanVar()
    self.checkbutton_open = Checkbutton(tab_main, text='Open document', variable=self.bool_open_document, onvalue=True, offvalue=False, font=font_1, borderwidth=3)
    self.checkbutton_open.grid(row=2, column=2, sticky=E)
    self.checkbutton_open.select()
    self.button_fetch = Button(tab_main, text='Fetch', command=self.start_fetch_document, font=font_1, borderwidth=3)
    self.button_fetch.grid(row=2, column=3, sticky=E)
    self.status_listbox = Listbox(tab_main, font=font_1, height=5, borderwidth=3, highlightthickness=0, selectmode=SINGLE, activestyle=NONE)
    self.status_listbox.grid(row=3, column=1, columnspan=3, sticky=W+E, pady=(10, 0))
    for i, w in enumerate([0, 10, 2, 1, 5]):
      tab_main.columnconfigure(i, weight=w)
    self.tk.bind('<Return>', lambda event: self.start_fetch_document())

    # configuration tab
    label = Label(tab_configure, text='Documents folder:', font=font_2)
    label.grid(row=1, column=1, sticky=W)
    self.documents_entry = Entry(tab_configure, font=font_2, borderwidth=3)
    self.documents_entry.grid(row=1, column=2, sticky=W+E)
    self.documents_entry.insert(0, self.documents_path)
    button = Button(tab_configure, text='Select', command=self.select_documents_path, font=font_2, borderwidth=3, anchor=W)
    button.grid(row=1, column=3, sticky='')
    label = Label(tab_configure, text='BDSS status', font=font_2)
    label.grid(row=2, column=1, sticky=W)
    self.label_bdss_status = Label(tab_configure, font=font_2, borderwidth=3, anchor=W, relief='sunken', text='Checking')
    self.label_bdss_status.grid(row=2, column=2, sticky=W+E)
    label = Label(tab_configure, text='Script status', font=font_2)
    label.grid(row=3, column=1, sticky=W)
    self.button_script_status = Label(tab_configure, font=font_2, borderwidth=3, anchor=W, relief='sunken', text='Checking')
    self.button_script_status.grid(row=3, column=2, sticky=W+E)
    self.button_script_fetch = Button(tab_configure, text='Fetch', command=self.start_fetch_script, font=font_2, borderwidth=3, anchor=W, state=DISABLED)
    self.button_script_fetch.grid(row=3, column=3, sticky='')
    label = Label(tab_configure, text='Database status', font=font_2)
    label.grid(row=4, column=1, sticky=W)
    self.button_database_status = Label(tab_configure, font=font_2, borderwidth=3, anchor=W, relief='sunken', text='Checking')
    self.button_database_status.grid(row=4, column=2, sticky=W+E)
    button = Button(tab_configure, text='Visit usptodata.com', command=self.open_website, font=font_2, borderwidth=3, anchor=CENTER, state=NORMAL)
    button.grid(row=5, column=1, pady=10, sticky=W+E)
    label_text = ' for more information about the BDSS TAR Index project.'
    Label(tab_configure, text=label_text, font=font_2).grid(row=5, column=2, columnspan=2, sticky=W)
    for i, w in enumerate([1, 1, 8, 1]):
      tab_configure.columnconfigure(i, weight=w)

  def check_output_queue(self):
    """ Receives responses from WorkerProcess via output queue. """

    try:
      command, output_text, done = self.output_queue.get_nowait()
      if command == 'bdss_status':
        self.label_bdss_status.config(text=output_text)
        if output_text != 'Online':
          self.update_status('Warning: BDSS is offline. No files can be retrieved.')
      elif command == 'script_status':
        self.button_script_status.config(text=output_text)
        self.button_script_fetch['state'] = DISABLED if output_text == 'OK' else NORMAL
      elif command == 'db_status':
        self.button_database_status.config(text=output_text)
      elif command == 'fetch_script':
        self.button_script_status.config(text=output_text)
        self.button_script_fetch['state'] = DISABLED if output_text == 'OK' else NORMAL
      elif command == 'status':
        self.update_status(output_text)
      elif command == 'fetch':
        self.update_status(output_text)
        self.button_fetch['state'] = DISABLED if output_text.startswith('Downloading') else NORMAL
      if done is True and command in self.workers:
        del self.workers[command]
    except queue.Empty:
      pass
    self.tk.after(100, self.check_output_queue)

  # event handler functions

  def select_documents_path(self):
    """ Allows user to select a documents path. """

    documents_path = filedialog.askdirectory(initialdir=self.documents_path,
      title='Documents Path')
    if os.path.isdir(documents_path) is False:
      return
    self.documents_path = documents_path
    self.documents_entry.delete(0, END)
    self.documents_entry.insert(0, documents_path)
    self.write_config()

  def start_fetch_document(self):
    """ Sends worker process a command to fetch a document. """
    if 'fetch' in self.workers:    # don't run two fetches at once
      return
    self.button_fetch['state'] = DISABLED
    command_arguments = (self.entry.get(), self.documents_path, self.bool_open_document.get())
    self.start_worker('fetch', command_arguments)

  def start_fetch_script(self):
    """ Sends worker process a command to fetch the latest bti.py script. """
    self.button_script_fetch['state'] = DISABLED
    self.start_worker('fetch_script')

  def open_website(self):
    """ Opens a web browser to the information page for BTI at usptodata.com. """
    webbrowser.open('https://www.usptodata.com/bti_page')

  # utility functions

  def update_status(self, s):
    self.status_listbox.insert('0', s)

  def start_worker(self, command, command_arguments = None):
    if command in self.workers:  # can't start two commands at the same time
      return
    worker = multiprocessing.Process(target=WorkerProcess.run, args=(command, command_arguments, self.output_queue))
    worker.daemon = True
    worker.start()
    self.workers[command] = worker

  def write_config(self):
    config = {'documents_path': self.documents_path}
    with open('bti_config.txt', 'wt', encoding='UTF-8') as f:
      f.write(json.dumps(config))

if __name__ == '__main__':
  BDSS_Interface.run_from_command_line()
