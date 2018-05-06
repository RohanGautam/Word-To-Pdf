import sys
import os
import comtypes.client

wdFormatPDF = 17

class locationError(Exception):pass # raised when location is not found

try:
    in_file = os.path.abspath(sys.argv[1])    
    out_file = os.path.abspath(sys.argv[2]) # saves in the same location as the input file
    if not os.path.exists(in_file): # If the file doesn't exist
        raise locationError
    #the conversion process:    
    print('Creating pdf...')
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()
    print('pdf created!')

#Errors that might be encountered:
except IndexError:
    print ('\nEnter the file names in the correct format:\n\tpython to-pdf.py <name/path of word file> <name of output file>')
except locationError:
    print ('Incorrect path/name of input file')
