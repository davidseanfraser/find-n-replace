import flask, os, shutil, datetime, zipfile, openpyxl, time
from werkzeug.utils import secure_filename

app = flask.Flask(__name__)
app.config['SECRET_KEY'] = 'asdofijasdfoijasdfoij'
app.config['UPLOAD_FOLDER'] = r'C:\Users\David Fraser\Desktop\CS\fnr\UPLOAD_FOLDER'

OLD_FILE_DURATION = 300

@app.route('/',methods=['POST', 'GET'])
def main():
    if flask.request.method == 'POST':
        if 'source_target_upload' not in flask.request.files:
            flask.flash('No file')
            return flask.redirect(flask.request.url)
        
        source_target_file = flask.request.files['source_target_upload']
        mapped_file = flask.request.files['mapped_file_upload']

        flask.request.files      
        source_target_filename = check_save_file(source_target_file)
        map_filename = check_save_file(mapped_file)
        
        dir_src_tgt = os.path.join(app.config['UPLOAD_FOLDER'], source_target_filename)
        dir_map_file = os.path.join(app.config['UPLOAD_FOLDER'], map_filename)

        ret_filename = director(dir_map_file, dir_src_tgt)
        
        return flask.redirect(flask.url_for('download_file', filename=ret_filename))
        
    return flask.render_template('index.html')

def delete_old_files():
    path = app.config['UPLOAD_FOLDER']
    files = os.listdir(path)
    cur_time = time.time()
    
    for file in files:
        name, ext = os.path.splitext(file)
        if ext != '':
            if (cur_time - os.stat(path+ '\\' + file).st_mtime) > OLD_FILE_DURATION:
                os.remove(path+'\\'+file)
   

delete_old_files()
def check_save_file(file):
    if file.filename == '':
        return False
    filename = secure_filename(file.filename)
    file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
    return filename

""" Development work - attempt to create file in memory to return whilst 
deleting the files in the path.  Works with non zip files.
@app.route('/uploads2/<map_file>/<src_file>')
def download_file2(map_file, src_file):
    map_path = os.path.join(app.config['UPLOAD_FOLDER'], map_file)
    src_path = os.path.join(app.config['UPLOAD_FOLDER'], src_file)
    
    def generate():
        with open(map_path) as f:
            yield from f
        
        os.remove(map_path)
        os.remove(src_path)
    
    r = app.response_class(generate(), mimetype = 'zip/text/csv') #, mimetype='text/csv'
    r.headers.set('Content-Disposition', 'attachment', filename='data.zip')
    return r
"""

@app.route('/uploads/<filename>')
def download_file(filename):
    delete_old_files()
    return(flask.send_from_directory(app.config['UPLOAD_FOLDER'],filename))
    
def director(file, *argv):
    name, ext = os.path.splitext(file)
    if len(argv) == 2:
        if ext == ".zip":
            unpack_map_remove(argv[0],argv[1], file)
        else:
            replace_file(argv[0],argv[1], file)
            
    if len(argv) == 1:
        if ext == ".zip":
            return map_zip_remove(argv[0], file)
        else:
            maps = map_file(argv[0])
            return replace_file(maps[0],maps[1], file)

def map_file(file):
    """ 

    Reads excel file, returns list of lists, column 1 being the first list
    of oldArr and 2nd column being the second list.  Sorts the list by the 
    first list so shorter matches won't occur first when longer possibilities
    exist
    
    Args:
        file: excel file that has two columns of string data of equal length on
              the first sheet
              
    Returns:
        list of two lists, both are lists of strings of equal length, sorted
        for length so longer values of the first list appear first
        
    Raises:
        ValueError: arrays are of unequal length
    """
    wb = openpyxl.load_workbook(filename=file)
    ws = wb.active
    split_array = []
    for col in ws.iter_cols():
        cur_array = []
        for cell in col:
            if str(cell.value) != 'None':
                cur_array.append(str(cell.value))
        split_array.append(cur_array)
        
    if len(split_array[0]) != len(split_array[1]):
        raise ValueError('Source and Target lists are not of equal length')
    reverse_sort = sorted([*zip(split_array[0],split_array[1])], key=lambda x: len(x[0]), reverse=True)
    unzipped_array = map(list,(zip(*reverse_sort)))
    _ = [*unzipped_array]
    
    return(_)

def map_zip_remove(mapfile, zipfile):
    """ Processes mapfile and passes values as arguments to mapping function"""
    
    arg_array = map_file(mapfile)
    return unpack_map_remove(arg_array[0],arg_array[1],zipfile)

def unpack_map_remove(oldArr, newArr, zipfile):
    """ Unpacks zip file, runs mapping process, rezips output, deletes folder
    
    Creates zip file that is processed using oldArr and newArr against a zip
    and tidies up all processing.
    
    Args:
        file: zip file to be processed, can contain arbitrary sub-dirs,
              all files are mapped
              
    Returns:
        Absolute directory of the zip archive that was created
        
    Raises:
        None
    """
    
    newpath = unzip_file(zipfile)
    
    new_zip_dir = map_zip_directory(oldArr, newArr, newpath)
    new_zip_filename = os.path.basename(new_zip_dir)
    shutil.rmtree(newpath)
    
    return new_zip_filename
    
def unzip_file(file):
    """ Unpacks zip file into folder of the same name sans file extension 
    
    Returns:
        A string representation of absolute directory of the path of unpacked
        file
    """
    path = str(file)[:-4]
    with zipfile.ZipFile(file) as myzip:
        myzip.extractall(path)
    return(path)
    
def map_zip_directory(oldArr,newArr,directory):
    """ Creates copy of directory, runs mapping process on new files, archives
    """
    new_directory = copy_directory(directory)
    
    replace_directory(oldArr,newArr,new_directory)
    
    shutil.make_archive(new_directory, 'zip', new_directory)
    
    shutil.rmtree(new_directory)
    
    return(str(new_directory)+'.zip')
    
def copy_directory(directory):
    """ Copies directory contents and creates new files with timestamped folder
    
    Args:
        directory: a directory which contents will be copied
        
    Returns:
        A string that is the absolute directory with timestamp of new folder
    
    Raises:
        None
    """
    
    timestamp = datetime.datetime.today().strftime("%Y-%m-%d-%H%M%S")
    new_name = directory + "_" + timestamp
    shutil.copytree(directory, new_name)
    return(new_name)
    

            
def replace_directory(oldArr, newArr, directory):
    """ Runs utility to replace oldArr values with newArr for all files in dir
    
    Args:
        oldArr: array of strings to be replaced by same index of newArr
        newArr: arry of strings to substitue in for oldArr values
        directory: string of directory that will be processed in place
    
    Returns:
        None
    
    Raises:
        None
    """
    
    file_paths = absolute_file_paths(directory)

    for file in file_paths:
        replace_file(oldArr, newArr, file)
        
def absolute_file_paths(directory):
    """ Yields generator for string of absolute directory of all files in dir
    
    Args:
        directory: Directory to be walked through and all files collected
    
    Returns:
        generator that yields string of absolute directory of all files in
            passed in directory
        
    Raises:
        None
    """
    
    for dirpath,_,filenames in os.walk(directory):
        for f in filenames:
            yield os.path.abspath(os.path.join(dirpath, f))

def replace_files(oldArr, newArr, files):
    for file in files:
        replace_file(oldArr, newArr, file)

def replace_file(oldArr, newArr, file):
    """ Replaces contents of file with mapped contents in place.
    
    If filetype is Excel, runs replace_excel() as helper
    Opens the file and reads contents to pass  function 
    replace_multiple_values().
    Overwrites existing content with mapped string and closes file.
    
    
    Args:
        oldArr: Array of strings that will be replaced
        newArr: Array of strings that are substituting oldArr strings
        file: File to be mapped
    
    Returns:
        String value of the absolute directory of the file that has been mapped
    
    Raises:
        None
    """
    name, ext = os.path.splitext(file)
    if ext in ['.xlsx', '.xlsm', '.xslt', '.xltm']:
        replace_excel(oldArr, newArr, file)
    else:
        f = open(file, 'r')
        filedata = f.read()
        f.close()
        
        writedata = replace_multiple_values(oldArr, newArr, filedata)
        
        f = open(file, 'w')
        f.write(writedata)
        f.close()
    
    return (os.path.basename(file))
        

def replace_excel(oldArr, newArr, file):
    """ Processes all sheets of excel, replacing strings
    
    iterates through each sheet, amassing list of cells through column iter,
    accesses each cell and will run through each mapping for each cell
    
    Args:
        file: Excel file passed for processing
        
    Returns:
        Saves Excel file in place with new mapped values
    
    Raises:
        None
    """
    wb = openpyxl.load_workbook(filename=file)
    for sheet in wb:
        for col in sheet.iter_cols():
            for cell in col:
                for i in range(len(oldArr)):
                    if type(cell.value) is str:
                        cell.value = cell.value.replace(oldArr[i], newArr[i])

    wb.save(file)
        

def replace_multiple_values(oldArr, newArr, filedata):
    """Replaces values in oldArr with the same index of newArr in the filedata
    
    Iterates over each value in oldArr and replaces it with matching index in 
    newArr.  
    
    Args:
        oldArr: Array of strings that will be replaced
        newArr: Array of strings that are substituting oldArr strings
        filedata: Already read data from a file
    
    Returns:
        A string of data that has all values of oldArr replaced
    
    Raises:
       Raise error when len(oldArr) != len(newArr)
    """ 
    # TODO: Should raise issue when len(oldArr) != len(newArr)
    newdata = filedata
    for i in range(len(oldArr)):
        newdata = newdata.replace(oldArr[i], newArr[i])
    return newdata


def print_output(file):
    """Prints each line in file.
    
    Opens file and prints each line before closing.
    
    Args: 
        file: A file to be parsed and printed
    
    Return:
        None
        
    """
    
    f = open(file, 'r')
    for line in f:
        print(line)
    f.close()


def read_excel(file):
    """ Quickly accesses and prints Excel values, for testing purposes"""
    wb = openpyxl.load_workbook(filename=file)
    for sheet in wb:
        for col in sheet.iter_cols():
            for cell in col:
                print(cell.value)

