from xml.etree.ElementTree import ElementTree
from xml.etree.ElementTree import Element
import xml.etree.ElementTree as etree
from xml.dom import minidom

import openpyxl
global_data = ''
global_data_initial_val=0
Bit_length_val=0
tc_incement_varible=0
max_value=0
Mid_value=0


def Access_signals():
    global tc_incement_varible
    # Testgroup issubfuction of root
    testgroup = etree.SubElement(root, "testgroup")
    testgroup.set('title', 'CAN Rx Signal '+ global_data + ' Test Using XCP')

    # testcase issubfuction of Testgroup
    testcase = etree.SubElement(testgroup, "testcase")
    testcase.set('ident', 'tc'+str(tc_incement_varible))
    testcase.set('title', 'Check XCP Parameter ' + global_data + ' update')

    # initialize is subfunction of testcase
    initialize = etree.SubElement(testcase, "initialize")
    initialize.set('title', 'Intialize parameters')
    initialize.set('wait', '500')

    # envvar is subfunction of initialize
    envvar = etree.SubElement(initialize, "envvar")
    envvar.set('name', 'ENV_' + global_data + '_IVC_R')
    envvar.text = str(''+str(global_data_initial_val)+'')  #intial value for input variable

    # Set is subfunction of testcase
    set = etree.SubElement(testcase, "set")
    set.set('title', 'Set')

    envvar = etree.SubElement(set, "envvar")
    envvar.set('name', 'ENV_' + global_data)
    envvar.text = ' 1 '

    # wait is subfunction of testcase
    wait = etree.SubElement(testcase, "wait")
    wait.set('title', 'Waiting for 1 Second')
    wait.set('time', '1000')

    # statecheck is subfunction of testcase
    statecheck = etree.SubElement(testcase, "statecheck")
    statecheck.set('wait', '1000')
    statecheck.set('title', 'State Check')

    # expected is subfunction of statecheck
    expected = etree.SubElement(statecheck, "expected")

    # envvar is subfunction of expected
    envvar = etree.SubElement(expected, "envvar")
    envvar.set('name', 'ENV_' + global_data + '_IVC_R')
    envvar.text = ' 1 '

    ##########################second test case statrt here ######################################
    tc_incement_varible=tc_incement_varible+1 #incement tc variable
    # testcase issubfuction of Testgroup
    testcase = etree.SubElement(testgroup, "testcase")
    testcase.set('ident', 'tc'+str(tc_incement_varible))
    testcase.set('title', 'Check XCP Parameter ' + global_data + ' Max Value')

    # Set is subfunction of testcase
    set = etree.SubElement(testcase, "set")
    set.set('title', 'Set')

    # envvar is subfunction of Set
    envvar = etree.SubElement(set, "envvar")
    envvar.set('name', 'ENV_' + global_data)
    envvar.text = str(max_value)

    # wait is subfunction of testcase
    wait = etree.SubElement(testcase, "wait")
    wait.set('title', 'Waiting for 1 Second')
    wait.set('time', '1000')
    #
    # statecheck is subfunction of testcase
    statecheck = etree.SubElement(testcase, "statecheck")
    statecheck.set('wait', '1000')
    statecheck.set('title', 'State Check')

    # expected is subfunction of statecheck
    expected = etree.SubElement(statecheck, "expected")

    # envvar is subfunction of expected
    envvar = etree.SubElement(expected, "envvar")
    envvar.set('name', 'ENV_' + global_data + '_IVC_R')
    envvar.text = str(max_value)


##########################Mid test case statrt here ######################################
    tc_incement_varible=tc_incement_varible+1 #incement tc variable
    # testcase issubfuction of Testgroup
    testcase = etree.SubElement(testgroup, "testcase")
    testcase.set('ident', 'tc'+str(tc_incement_varible))
    testcase.set('title', 'Check XCP Parameter ' + global_data + ' Mid Value')

    # Set is subfunction of testcase
    set = etree.SubElement(testcase, "set")
    set.set('title', 'Set')

    # envvar is subfunction of Set
    envvar = etree.SubElement(set, "envvar")
    envvar.set('name', 'ENV_' + global_data)
    envvar.text = str(Mid_value)

    # wait is subfunction of testcase
    wait = etree.SubElement(testcase, "wait")
    wait.set('title', 'Waiting for 1 Second')
    wait.set('time', '1000')
    #
    # statecheck is subfunction of testcase
    statecheck = etree.SubElement(testcase, "statecheck")
    statecheck.set('wait', '1000')
    statecheck.set('title', 'State Check')

    # expected is subfunction of statecheck
    expected = etree.SubElement(statecheck, "expected")

    # envvar is subfunction of expected
    envvar = etree.SubElement(expected, "envvar")
    envvar.set('name', 'ENV_' + global_data + '_IVC_R')
    envvar.text = str(Mid_value)



def readxls_file():
    global global_data
    global global_data_initial_val
    global Bit_length_val
    global tc_incement_varible
    global max_value
    global Mid_value
    wb = openpyxl.load_workbook('170510_ADAS_DRV_ARXML_IF rev8 VB.xlsx')
    type(wb)
    sheet=wb.get_sheet_by_name("ADAS 통합제어")
    for i in range(5,45):
     global_data = (sheet['J'+str(i)].value)
     local_data_initial_val=str((sheet['G'+str(i)].value))# conversion from hex to int
     global_data_initial_val=int(local_data_initial_val,16)
     Bit_length_val = str((sheet['E' + str(i)].value))
     print(global_data)
     tc_incement_varible=tc_incement_varible+1
     max_value_local=2**int(Bit_length_val)
     max_value=max_value_local-1
     Mid_value = int((max_value + global_data_initial_val) / 2)
     Access_signals()

    #sheet = book.active
    return


# def max_global_data_val():
#     global Bit_length_val
#     global max_value
#     max_value=2**Bit_length_val
#     print(max_value)
#     return

#def prettify(elem):
#    """Return a pretty-printed XML string for the Element.
#    """
#    rough_string = etree.tostring(elem, 'iso-8859-1')
#    reparsed = minidom.parseString(rough_string)
#    return reparsed.toprettyxml(indent="  ")


def indent(elem, level=0):
    i = "\n" + level*"  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level+1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i



root=Element('testmodule')

rem = etree.Comment("<!-Vector Test Automation Editor 2.1.34.0->")#added comment
root.append(rem)
rem1 = etree.Comment("<!-Version: 1.0->")#added comment
root.append(rem1)

tree=ElementTree(root)
root.set('title','XCP Demo Testmodule')
root.set('version','1.0')
root.set('xmlns','http://www.vector-informatik.de/CANoe/TestModule/1.16')

readxls_file()

#print(prettify(root))
#output=prettify(root)
#print(etree.tostring(root))
#tree.write("person.xml")
indent(root)
# writing xml


tree.write("person.xml", encoding="utf-8", xml_declaration=True)

#tree.write('person.xml','w',pretty_print=True)
