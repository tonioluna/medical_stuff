import os
import time
import sys
import logging
import traceback
import re
import csv
import colorama
import codecs
try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None

if __name__ == "__main__":
    I_AM_SCRIPT = True
    _me = os.path.splitext(os.path.basename(sys.argv[0]))[0]
    _my_path = os.path.dirname(sys.argv[0])
else:
    I_AM_SCRIPT = False
    _my_path = os.path.dirname(__file__)
_my_path = os.path.realpath(_my_path)

log = None

_true_values = ("true", "yes", "1", "sure", "why_not")
_false_values = ("false", "no", "0", "nope", "no_way")
_valid_input_file_exts = (".txt",)

def init_logger():
    global log
    
    log = logging.getLogger(_me)
    log.setLevel(logging.DEBUG)
    
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    filename = _me + ".log"
    with open(filename, "w") as fh:
        fh.write("Starting on %s running from %s"%(time.ctime(), repr(sys.argv)))
    fh = logging.FileHandler(filename = filename)
    fh.setLevel(logging.DEBUG)
    
    
    # create formatter
    formatter = logging.Formatter('%(levelname)s - %(message)s')
    
    ch.setFormatter(formatter)
    fh.setFormatter(formatter)
    
    log.addHandler(ch)
    log.addHandler(fh)
    
    colorama.init()

    log = LoggerWrapper(log)
    
    log.info("Started on %s at %s running under %s"%(time.ctime(), 
                                                     os.environ.get("COMPUTERNAME", "UNKNOWN_COMPUTER"),
                                                     os.environ.get("USERNAME", "UNKNOWN_USER"),))

class LoggerWrapper:
    
    def __init__(self, parent_logger):
        self._parent_logger = parent_logger

    def __getattr__(self, key):
        if key == "info":
            sys.stdout.write(colorama.Fore.RESET)
        elif key == "debug":
            sys.stdout.write(colorama.Fore.MAGENTA)
        elif key == "warning":
            sys.stdout.write(colorama.Fore.YELLOW)
        elif key == "error":
            sys.stdout.write(colorama.Fore.RED)
        return getattr(self._parent_logger, key)


def find_input_files():
    dir = os.path.realpath(os.getcwd())
    log.info("Looking for files on %s"%(dir,))
    items = os.listdir(dir)
    items.sort()
    files = []
    for i in items:
        if not os.path.splitext(i)[1].lower() in _valid_input_file_exts: continue
        files.append(os.path.join(dir, i))
    return files
            
_re_line_with_description = re.compile("^(?P<parameter>.*)\s(?P<curr_val>\S*)\s(?P<units>(Amarillo|Claro|Negativo|Ausentes).*)")
_re_line_less_than = re.compile("^(?P<parameter>.*)\s(?P<curr_val>\d[.\d]*)\s\<[\s]*(?P<max_val>\d[.\d]*)[\s]*(?P<units>.*)")
_re_line_more_than = re.compile("^(?P<parameter>.*)\s(?P<curr_val>\d[.\d]*)\s\>[\s]*(?P<min_val>\d[.\d]*)[\s]*(?P<units>.*)")
#_re_line_less_than = re.compile("^(?P<parameter>.*)\s(?P<curr_val>\d[.\d]*)\s\<\s(?P<max_val>\d[.\d]*)[\s]*(?P<units>.*)")
_re_line_with_range = re.compile("^(?P<parameter>.*)\s(?P<curr_val>\d[.\d]*)\s(?P<min_val>\d[.\d]*)[\s]*\-[\s]*(?P<max_val>\d[.\d]*)[\s]*(?P<units>.*)")

           
_test_examen_orina = ("EXAMEN GENERAL DE ORINA", "Orina.")
_test_quimica_45 = ("QUÍMICA INTEGRAL DE 45 ELEMENTOS", "Sangre.")
_test_quimica_6 = ("QUÍMICA 6", "Sangre.")
_test_hemog_glicos = ("HEMOGLOBINA GLICOSILADA A1c", "Sangre.")
_test_acido_urico = ("ÁCIDO ÚRICO EN SANGRE", "Sangre.")

_tests = (_test_examen_orina,
          _test_quimica_45,
          _test_quimica_6,
          _test_hemog_glicos,
          _test_acido_urico,
         )
           
_line_regexes = [_re_line_with_description,
                 _re_line_with_range,
                 _re_line_less_than,
                 _re_line_more_than,
                ]
           
def _set_dict_defaults(d):
    if "min_val" not in d: d["min_val"] = None
    if "max_val" not in d: d["max_val"] = None
    if "units" not in d: d["units"] = None
           
def get_file_entries(file):
    log.info("Reading %s"%(file,))
    #with open(file) as fh:
    results = []
    with codecs.open(file, "r", "utf8") as fh:    
        curr_prefix = None
        for line in fh:
            line = line.replace("\r","").replace("\n","")
            
            # Match a test name
            for test in _tests:
                if line == test[0]:
                    curr_prefix = test[1]
                    break

    
            # Match a regex
            found_match = False
            for regex in _line_regexes:
                m = regex.match(line)
                if m != None:
                    d = m.groupdict()
                    
                    if curr_prefix == None:
                        log.warning("Found result wihtout prefix ready")
                        curr_prefix = ""
                    d["parameter"] = curr_prefix + d["parameter"]
                    
                    _set_dict_defaults(d)
                    results.append(d)
                    log.debug("LINE OK        : %s"%(line,))
                    found_match = True
            if not found_match:
                log.debug("LINE INVALID   : %s"%(line,))
    
    return results     
       
def merge_file_entries(entries):
    merged = {}
    # Phase 1, create merge dictionary
    for filename, entries in entries.items():
        for entry in entries:
            parameter = entry["parameter"]
            if not parameter in merged:
                merged[parameter] = {}
                
            merged[parameter][filename] = entry
    
    # Phase 2, check all entries have same max, min and units
    parameter_common_items = {}
    for test in ("min_val", "max_val", "units"):
        for parameter, file_dict in merged.items():
            
            if parameter not in parameter_common_items:
                parameter_common_items[parameter] = {}
            
            found_values = []
            for filename, entry_dict in file_dict.items():
                v = entry_dict[test]
                if v not in found_values:
                    found_values.append(v)
            if len(found_values) != 1:
                log.error("Multiple values for test %s on parameter %s: %s"%(test, repr(parameter), ", ".join([repr(s) for s in found_values])))
                parameter_common_items[parameter][test] = "## MULTIPLE ##"
            else:
                parameter_common_items[parameter][test] = found_values[0]
    
    # Phase 3, build the actual merged dictionary
    results = {}
    for parameter, file_dict in merged.items():
        results[parameter] = {}
        results[parameter]["__min_val__"] = parameter_common_items[parameter]["min_val"]
        results[parameter]["__max_val__"] = parameter_common_items[parameter]["max_val"]
        results[parameter]["__units__"] = parameter_common_items[parameter]["units"]
        
        for filename, entry_dict in file_dict.items():
            results[parameter][filename] = merged[parameter][filename]["curr_val"]
    
    return results
    
def write_csv(results, keys, report_name):
    filename = report_name + ".csv"
    log.info("Writting report to %s"%(filename,))
    with open(filename, "w", newline="") as fh:
        writer = csv.writer(fh)
        hdr = ["Parametro", "Min Val", "Max Val", "Units"]
        for k in keys:
            hdr.append(os.path.splitext(os.path.basename(k))[0].replace("Antonio Luna","").strip())
        writer.writerow(hdr)
        parameters = list(results.keys())
        parameters.sort()
        for parameter in parameters:
            d = results[parameter]
            row = []
            #print("START")
            row.append(parameter)
            row.append(d["__min_val__"])
            row.append(d["__max_val__"])
            row.append(d["__units__"])
            #print(len(row))
            for k in keys:
                v = d.get(k, ".")
                #print(k, v)
                row.append(v)
                #print(len(row))
            #print(repr(row))
            #print(len(row))
            writer.writerow(row)
            #assert 1 == 0
        
def write_xlsx(results, keys, report_name):
    assert xlsxwriter != None, "xlswriter module is missing, cannot generate XLSX format"

    filename = report_name + ".xlsx"
    log.info("Writting report to %s"%(filename,))
    
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    cell_format_red_text = workbook.add_format({'font_color': 'red'})

    # Start from the first cell. Rows and columns are zero indexed.
    rnum = 0
    col = 0

    hdr = ["Parametro", "Min Val", "Max Val", "Units"]
    for k in keys:
        hdr.append(os.path.splitext(os.path.basename(k))[0].replace("Antonio Luna","").strip())
    
    for index, cell in enumerate(hdr):
        worksheet.write(rnum, index, cell)
    
    rnum += 1
    
    parameters = list(results.keys())
    parameters.sort()
    for parameter in parameters:
        d = results[parameter]
        row = []
        #print("START")
        row.append((parameter, None))
        min_val = try_num(d["__min_val__"])
        max_val = try_num(d["__max_val__"])
        row.append((min_val, None))
        row.append((max_val, None))
        row.append((d["__units__"], None))
        #print(len(row))
        for k in keys:
            v = try_num(d.get(k, "."))
            use_red = False
            if type(v) == float:
                if type(min_val) == float:
                    use_red = use_red or v < min_val
                if type(max_val) == float:
                    use_red = use_red or v > max_val
            
            row.append((v, None if not use_red else cell_format_red_text))
        for index, cell in enumerate(row):
            worksheet.write(rnum, index, cell[0], cell[1])
        rnum += 1
    
    
    workbook.close()

def try_num(txt):
    try:
        return float(txt)
    except:
        return txt

def main():
    init_logger()
    
    try:
        files = find_input_files()
        
        file_entries = {}
        for file in files:
            file_entries[file] = get_file_entries(file)
        
        merged_entries = merge_file_entries(file_entries)
        
        filename = time.strftime("merged_%y%m%d_%H%M%S")
        write_csv(merged_entries, files, report_name = filename)
        write_xlsx(merged_entries, files, report_name = filename)
        
    except Exception as ex:
        log.error("Caught high-level exception: %s"%(ex,))
        log.debug(traceback.format_exc())
        return 1
    return 0
    
if I_AM_SCRIPT:
    rc = main()
    if type(rc) is not int:
        rc = -1
    exit(rc)
