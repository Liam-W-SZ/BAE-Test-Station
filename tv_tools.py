import jsonpickle
import ftplib
import os
import time
import json
# import tb_rest_client
import pycurl
from io import BytesIO,StringIO
import re
#from cStringIO import StringIO
import shutil  # Add this import for file operations



class root:
  def __init__(self, ts, tester, testdesc, testjig, deviceid, prefix, jobnr, szserialnr, procedure, productGroup, supplierserial, lowerlevelID, tests, errors, result):
    self.ts = ts            
    self.tester = tester      
    self.testdesc = testdesc
    self.testjig = testjig
    self.deviceid = deviceid
    self.prefix = prefix
    self.jobnr = jobnr
    self.szserialnr = szserialnr
    self.procedure = procedure
    self.productGroup = productGroup
    self.supplierserial = supplierserial
    self.lowerlevelID = lowerlevelID
    self.tests = tests #list of tests
    self.errors = errors
    self.result = result
   
def testResultReset (root):
    root.ts = ""            
    root.tester = root.tester      
    root.testdesc = ""
    root.testjig = ""
    root.deviceid = ""
    root.jobnr = ""
    root.szserialnr = ""
    root.procedure = "ATP"
    root.supplierserial = ""
    root.lowerlevelID = ""
    root.tests = []
    root.result = ""
    return root

class test:
  def __init__(self, description, expected, received, result, test_type, test_range, required):
    self.description = description
    self.expected = expected
    self.received = received
    self.result = result
    self.test_type = test_type
    self.test_range = test_range
    self.required = required
    
class Device:
  def __init__(self, product, szs, deviceName, tbDeviceID, tbDeviceLabel, sessionJWT):
    self.product = product
    self.szs = szs
    self.deviceName = deviceName
    self.tbDeviceID = tbDeviceID
    self.tbDeviceLabel = tbDeviceLabel
    self.sessionJWT = sessionJWT

def outputJSON(root, filepath, filename, ftpfolder):
    """
    Takes a JSON file stored locally and uploads it to the FTP server.
    """
    try:
        # Check if source file exists
        if not os.path.exists(filepath):
            print(f"Source file not found: {filepath}")
            return False

        # Upload to FTP server
        session = ftplib.FTP('13.245.219.204') #change from FTP_TLS to FTP. Remove f.prot_p()
        session.login('tvpuser','r136I5xS')
        #session.prot_p()
        session.cwd(ftpfolder)
        
        with open(filepath, 'rb') as file:
            session.storbinary(f'STOR {filename}', file)
        
        session.quit()
        print(f"Successfully uploaded {filename} to FTP server")
        return True

    except FileNotFoundError as e:
        print(f"File not found error: {e}")
    except ftplib.all_errors as e:
        print(f"FTP error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    
    return False

def outputEXCEL(excel_results_folder, file_name, ftpfolder, result):
    """
    Takes a file from excel_results_folder and uploads it to the FTP server.
    """
    try:


        # Get the new filename with result suffix
        base_name, ext = os.path.splitext(file_name)
        new_file_name = f"{base_name}{result}{ext}"

        # Check source file
        source_file = os.path.join(excel_results_folder, new_file_name)
        if not os.path.exists(source_file):
            print(f"Source file not found: {source_file}")
            return False


        # Upload to FTP server
        session = ftplib.FTP('13.245.219.204') #change from FTP_TLS to FTP. Remove f.prot_p()
        session.login('tvpuser','r136I5xS')
        #session.prot_p()
        session.cwd(ftpfolder)
        
        with open(source_file, 'rb') as file:
            session.storbinary(f'STOR {new_file_name}', file)
        
        session.quit()
        print(f"Successfully uploaded {new_file_name} to FTP server")
        return True

    except FileNotFoundError as e:
        print(f"File not found error: {e}")
    except ftplib.all_errors as e:
        print(f"FTP error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
    
    return False

def outputJSON_local(root,filepath):
    s = jsonpickle.encode(root, unpicklable = False)
    f = open(filepath, "w")  # Change mode to "w" to overwrite if file exists
    f.write(s)
    f.close()
    return s

def outputEXCEL_local(excel_results_folder, aging_data_folder, file_name, result):
    """
    Moves an Excel file from the Aging_Data folder to the Excel Results folder on the Desktop,
    renames it with test result suffix, and deletes the original file.
    """
    try:
        # Ensure the Excel Results folder exists
        if not os.path.exists(excel_results_folder):
            os.makedirs(excel_results_folder)

        # Check if source file exists
        source_file = os.path.join(aging_data_folder, file_name)
        if not os.path.exists(source_file):
            print(f"Source file not found: {source_file}")
            return

        # Get file extension and create new filename
        base_name, ext = os.path.splitext(file_name)
        new_file_name = f"{base_name}{result}{ext}"
        destination_file = os.path.join(excel_results_folder, new_file_name)

        # Copy file to destination
        shutil.copy2(source_file, destination_file)

        # Verify copy was successful
        if os.path.exists(destination_file):
            print(f"File copied successfully to: {destination_file}")
            # Delete original file only after successful copy
            try:
                os.remove(source_file)
                print(f"Original file deleted: {source_file}")
                return True
            except Exception as e:
                print(f"Warning: Could not delete original file: {e}")
                return False
        else:
            print(f"Failed to copy file to: {destination_file}")
            return False

    except FileNotFoundError as e:
        print(f"File not found error: {e}")
    except PermissionError as e:
        print(f"Permission error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")


def store_pack_info(ftp_folder,local_folder,filename):

    ftp_server = '13.245.219.204'
    ftp_user = 'tvpuser'
    ftp_pass = 'r136I5xS'

    # Connect to the FTP server securely
    session = ftplib.FTP(ftp_server)

    # Login to the server with username and password
    session.login(ftp_user, ftp_pass)

    #session.prot_p()

    # Change to the target directory on the FTP server
    session.cwd(ftp_folder)

    with open(local_folder, 'rb') as json_file:
        # Upload the JSON file 
        session.storbinary(f'STOR {filename}', json_file)

    # Close the FTP 
    session.quit()
    print(f"Successfully uploaded complete pack info to ftp server")
    #print(f"Successfully uploaded {local_folder} to {ftp_folder}/{filename}")

#Jasper Function
def curlPost(data_dict):
    curl = pycurl.Curl()
    curl.setopt(pycurl.URL, 'https://script.google.com/macros/s/AKfycbzsCDZ44HSq6S4tMZx_AOOK1DeXAjucrQ5o65AWBnVom0y5JG4i9-yXN_73qqh71mbW-A/exec')
    curl.setopt(pycurl.HTTPHEADER, ['Accept: application/json',
                                    'Content-Type: application/json'])
    curl.setopt(pycurl.POST, 1)

    # If you want to set a total timeout, say, 10 seconds
    curl.setopt(pycurl.TIMEOUT_MS, 10000)

    curl.setopt(pycurl.CAINFO, '/home/main/path/to/certs/certs/cacert.pem')  # Update this if needed

    ## depending on whether you want to print details on stdout, uncomment either
    # curl.setopt(pycurl.VERBOSE, 1) # to print entire request flow
    ## or
    # curl.setopt(pycurl.WRITEFUNCTION, lambda x: None) # to keep stdout clean

    # preparing body the way pycurl.READDATA wants it
    # NOTE: you may reuse curl object setup at this point
    #  if sending POST repeatedly to the url. It will reuse
    #  the connection.
    body_as_dict = {"name": "abc", "path": "def", "target": "ghi"}
    body_as_json_string = data_dict # dict to json
    body_as_file_object = StringIO(body_as_json_string)

    # prepare and send. See also: pycurl.READFUNCTION to pass function instead
    curl.setopt(pycurl.READDATA, body_as_file_object) 
    curl.setopt(pycurl.POSTFIELDSIZE, len(body_as_json_string))
    curl.perform()

    # you may want to check HTTP response code, e.g.
    status_code = curl.getinfo(pycurl.RESPONSE_CODE)
    if status_code != 200:
        print ("Aww Snap :( Server returned HTTP status code {}".format(status_code))

    # don't forget to release connection when finished

    print(f"Successfully uploaded complete test data to Jaspers Drive")
    curl.close()

#Jasper Function

def maskCheck(serial,ser_type):
    rex_R48 = re.compile(r"^[S]{1}[Z]{1}[S]{1}[:]{1}[S]{1}[Z]{1}[-]{1}[R]{1}[4]{1}[8]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[S]{1}[C]{1}[:]{1}[4]{1}[0]{1}[5]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[J]{1}[N]{1}[:]{1}[M]{1}[F]{1}[0-9]{6}")
    rex_R48S = re.compile(r"^[S]{1}[Z]{1}[S]{1}[:]{1}[S]{1}[Z]{1}[-]{1}[R]{1}[4]{1}[8]{1}[S]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[S]{1}[C]{1}[:]{1}[4]{1}[0]{1}[5]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[J]{1}[N]{1}[:]{1}[M]{1}[F]{1}[0-9]{6}")
    rex_B48S = re.compile(r"^[S]{1}[Z]{1}[S]{1}[:]{1}[S]{1}[Z]{1}[-]{1}[B]{1}[4]{1}[8]{1}[S]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[S]{1}[C]{1}[:]{1}[4]{1}[0]{1}[5]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[J]{1}[N]{1}[:]{1}[M]{1}[F]{1}[0-9]{6}")
    rex_V48 = re.compile(r"^[S]{1}[Z]{1}[S]{1}[:]{1}[S]{1}[Z]{1}[-]{1}[V]{1}[4]{1}[8]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[S]{1}[C]{1}[:]{1}[4]{1}[0]{1}[5]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[J]{1}[N]{1}[:]{1}[M]{1}[F]{1}[0-9]{6}")
    rex_H48 = re.compile(r"^[S]{1}[Z]{1}[S]{1}[:]{1}[S]{1}[Z]{1}[-]{1}[H]{1}[4]{1}[8]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[S]{1}[C]{1}[:]{1}[4]{1}[0]{1}[5]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[J]{1}[N]{1}[:]{1}[M]{1}[F]{1}[0-9]{6}")
    rex_VSA = re.compile(r"^[S]{1}[Z]{1}[S]{1}[:]{1}[S]{1}[Z]{1}[-]{1}[V]{1}[S]{1}[A]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[S]{1}[C]{1}[:]{1}[6]{1}[2]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[J]{1}[N]{1}[:]{1}[M]{1}[F]{1}[0-9]{6}")
    rex_HSA = re.compile(r"^[S]{1}[Z]{1}[S]{1}[:]{1}[S]{1}[Z]{1}[-]{1}[H]{1}[S]{1}[A]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[S]{1}[C]{1}[:]{1}[6]{1}[2]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[J]{1}[N]{1}[:]{1}[M]{1}[F]{1}[0-9]{6}")
    rex_C48 = re.compile(r"^[S]{1}[Z]{1}[S]{1}[:]{1}[S]{1}[Z]{1}[-]{1}[C]{1}[4]{1}[8]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[S]{1}[C]{1}[:]{1}[6]{1}[2]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[J]{1}[N]{1}[:]{1}[M]{1}[F]{1}[0-9]{6}")
    rex_C48S = re.compile(r"^[S]{1}[Z]{1}[S]{1}[:]{1}[S]{1}[Z]{1}[-]{1}[C]{1}[4]{1}[8]{1}[S]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[S]{1}[C]{1}[:]{1}[6]{1}[2]{1}[-]{1}[0-9]{7}\s{1}\/{1}\s{1}[J]{1}[N]{1}[:]{1}[M]{1}[F]{1}[0-9]{6}")


    
    if ser_type == "VSA":
        if rex_VSA.match(serial):
            #print("Correct format")
            return True
        else:
            #print("Incorrect Format, Rescan")
            #CTkMessagebox(title="Serial Number Error",message="Incorrect VSA/C48 Serial Number Entered. Please Rescan",icon="warning")                    
            return False
    
    if ser_type == "HSA":
        if rex_HSA.match(serial):
            #print("Correct format")
            return True
        else:
            #print("Incorrect Format, Rescan")
            #CTkMessagebox(title="Serial Number Error",message="Incorrect VSA/C48 Serial Number Entered. Please Rescan",icon="warning")                    
            return False
                        
    if ser_type == "C48":
        if rex_C48.match(serial):
            #print("Correct format")
            return True
            pass
        else:
            #print("Incorrect Format, Rescan")
            #CTkMessagebox(title="Serial Number Error",message="Incorrect VSA/C48 Serial Number Entered. Please Rescan",icon="warning")                    
            return False

    if ser_type == "C48S":
        if rex_C48S.match(serial):
            #print("Correct format")
            return True
            pass
        else:
            #print("Incorrect Format, Rescan")
            #CTkMessagebox(title="Serial Number Error",message="Incorrect VSA/C48 Serial Number Entered. Please Rescan",icon="warning")                    
            return False

    if ser_type == "V48":
        if rex_V48.match(serial):
            #print("Correct format")
            return True
        else:
            #CTkMessagebox(title="Serial Number Error",message="Incorrect 405 Serial Number Entered. Please Rescan",icon="warning")
            return False
    if ser_type == "R48":
        if rex_R48.match(serial):
            #print("Correct format")
            return True
        else:
            #CTkMessagebox(title="Serial Number Error",message="Incorrect 405 Serial Number Entered. Please Rescan",icon="warning")
            return False
    if ser_type == "R48S":
        if rex_R48S.match(serial):
            #print("Correct format")
            return True
        else:
            #CTkMessagebox(title="Serial Number Error",message="Incorrect 405 Serial Number Entered. Please Rescan",icon="warning")
            return False
    if ser_type == "B48S":
        if rex_B48S.match(serial):
            #print("Correct format")
            return True
        else:
            #CTkMessagebox(title="Serial Number Error",message="Incorrect 405 Serial Number Entered. Please Rescan",icon="warning")
            return False
    if ser_type == "H48":
        if rex_H48.match(serial):
            #print("Correct format")
            return True
        else:
            #CTkMessagebox(title="Serial Number Error",message="Incorrect 405 Serial Number Entered. Please Rescan",icon="warning")
            return False


def load_ftp_file(ftpfolder,filename):
        # Initialize the data variable
        file_data = b''  # Using 'b' for binary data storage
        session = ftplib.FTP_TLS('13.246.31.38')
        session.login('ftpuser','ftpuser')
        session.prot_p()
        session.cwd(ftpfolder)
        file_data = []
        # Define the callback function
        def file_callback(data):
            file_data.append(data)  # Append each chunk to the list
            #print(file_data)
        # Retrieve the file in binary mode
        session.retrbinary('RETR ' + filename, file_callback)
       
        # Combine the chunks into a single binary object
        file_data = b''.join(file_data)
        session.quit()

        return file_data
        # Load config from FTP Server


def getTelemetry(root, telemetryName):
  value = ""
  ts = 0
  return value, ts
  
def initTB():
    file_path = os.getcwd() +"/config/config.json"
    pempath = str(os.getcwd() +"/config/thingsboard_ca_bundle.pem")
    with open(file_path, 'r') as file:
        config = json.load(file)
    url = str(config.get("ThingsBoard").get("prod").get("url"))
    username = str(config.get("ThingsBoard").get("prod").get("username"))
    password = str(config.get("ThingsBoard").get("prod").get("password"))
    tb_Session = tb_rest_client.thingsboard(url, username, password, pempath)
    return tb_Session

def createTBDevice(tb_session, UUT):
    #print("In createTBDevice()")
    error = "None"
    file_path = os.getcwd() +"/config/config.json"
    with open(file_path, 'r') as file:
        config = json.load(file)
    mcaName = UUT.lowerlevelID.split('/')[0].split(':')[1].strip()
    deviceName = UUT.deviceid
    #print("MCA Name: " +mcaName)
    deviceIDresponse, error = tb_session.getDeviceID(mcaName)
    if error == "None":
        tb_session.device = {
                    "id": deviceIDresponse.get("id").get("id"),
                    "name": deviceIDresponse.get("name"),
                    "type": deviceIDresponse.get("type"),
                    "entityGroupId": None,
                    "profileId": deviceIDresponse.get("deviceProfileId").get("id"),
                    "label": deviceIDresponse.get("label")
                }
        tb_session.placeholder_device = tb_session.device
    else:
        deviceIDresponse, error = tb_session.getDeviceID(deviceName)
        if error == "None":
            tb_session.device = {
                        "id": deviceIDresponse.get("id").get("id"),
                        "name": deviceIDresponse.get("name"),
                        "type": deviceIDresponse.get("type"),
                        "entityGroupId": None,
                        "profileId": deviceIDresponse.get("deviceProfileId").get("id"),
                        "label": deviceIDresponse.get("label")
                    }
        else:
            error= "E:404"
    
    return tb_session, error
def AssignSystem(tb_session, UUT):
    #print("In AssignSystem()")
    if UUT.szserialnr not in tb_session.device.get("label"):
        file_path = os.getcwd() +"/config/config.json"
        with open(file_path, 'r') as file:
            config = json.load(file)
        System= UUT.prefix
        newID = tb_session.device.get("id")
        newName = UUT.deviceid
        success = False
        error = None
        newEntityGroupID = config.get("ThingsBoard").get("prod").get("products").get(System).get("group").get("id")
        newType= config.get("ThingsBoard").get("prod").get("products").get(System).get("profile").get("name")
        newProfileID = config.get("ThingsBoard").get("prod").get("products").get(System).get("profile").get("id")
        if UUT.prefix not in tb_session.device.get("label"):
            newLabel = tb_session.device.get("label") + "," + UUT.szserialnr
        else:
            newLabel = tb_session.device.get("label")
        newDevice = {
                        "id": newID,
                        "name": newName,
                        "type": newType,
                        "entityGroupId": newEntityGroupID,
                        "profileId": newProfileID,
                        "label": newLabel
                    }
        tb_session.device = newDevice
        #print("New Device:\n")
        #print(tb_session.device)
        success, error =tb_session.update_device()
        #print ("Success?:" + str(success))
        #print("Error: " + str(error))
        if success:
            results = "Success"
        else:
            results = "Failed"
        
    else:
        print("Device already assigned!")
        success = True
        results = "Success"
        error = None
    return success, error, results

def get_mca_atp_results(tb_session, UUT):
    results = ""
    att_string, error= tb_session.get_Attribute("atp_test_results")
    if error == None:
        success = True
        results=att_string[0].get("value")

    return success, error, results

def reset_tb_device(tb_session, UUT):
    tb_session.device = tb_session.placeholder_device
    success, error =tb_session.update_device()
    return success, error

def test_check(test):
    result = ""
    if test.test_range == "Binary":
        if test.received == test.expected:
            result = "Passed"
        else:
            result = "Failed"
    elif test.test_range == "Multi":
        if test.received in test.expected:
            result = "Passed"
        else:
            result = "Failed"
    return result

def set_rapidTelemetry(tb_session, switch):
    payload =tb_session.send_RPC("rapid_telemetry", switch)
    if payload == None:
        success = False
    elif payload.get("payload") == "OK":
        success = True
    else:
        success = False
    return success

def timeout_polling_ota(tb_session, next_state, time):
    currentstate = ""
    while currentstate != next_state:
        currentstate= tb_session.get_latest_time_series_value("fw_state").get("fw_state")[0].get("value")
        if retry == 0:
            return False, currentstate
        else:
            retry -= 1
    return True, currentstate    
def get_expected_firmware(UUT):

	if UUT.productGroup == "Genoa" or UUT.productGroup == "Faraday":
		firmware = "sz_mc1_prod_e4_modbus"
	elif UUT.productGroup == "Fatty":
		firmware = "sz_mc1_prod_e4_modbus_rot_disab"
	else:
		firmware = "sz_mc1_prod_tbb_rs485"
	return firmware
def check_alarms(check_list, UUT, config, alarmString):
    totalresult = True
    for x in check_list:
        test_list = config.get("Tests").get(x)
        newTest = test(test_list.get("Description"),test_list.get("expected"), "", "Untested",test_list.get("Test_Type"),test_list.get("Test_Range"),test_list.get("required"))
        if test_list.get("Alarm") in alarmString:
            newTest.received = "TRUE"
            UUT.errors.append(test_list.get("error"))
            totalresult = False
        else:
            newTest.received = "FALSE"
        newTest.result = test_check(newTest)
        UUT.tests.append(newTest)
    return totalresult
def get_config_files():
    filename = "config.json"
    A = os.getcwd() +"/config/" + filename
    print(A)
    session = ftplib.FTP_TLS('13.246.31.38')
    session.login('ftpuser','ftpuser')
    session.prot_p()
    session.cwd("Test_Software/System_ATP/config_files")
    try:
        session.retrbinary("RETR " + filename ,open(A, 'wb').write)
        print("Write successful")
    except:
        print("Error")
    session.quit



