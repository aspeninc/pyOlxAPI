"""
Prefault voltage profile factors check

To run inside of the OneLiner GUI:
1. Launch OneLiner
2. Open your network model (.olr file)
3. Run the Tools | Python OlxAPI Dashboard command within OneLiner
4. Select the PrefaultCheck.py script
5. Click the Launch button
6. A dialog to choose an output file for the CSV report will be displayed
7. The checking functions will be executed (may take several minutes for large network models)
8. A dialog will ask if you want to open the CSV report using your Windows default program

To run in Python or Python IDE.
1. Launch PrefaultCheck.py script in your Python (32-bit Python 3.9 or newer is required)
2. A dialog to choose a OneLiner .olr file will be displayed
3. A dialog to choose an output file for the CSV report will be displayed
4. The checking functions will be executed (may take several minutes for large network models)
5. A dialog will ask if you want to open the CSV report using your Windows default program
"""
__author__    = "ASPEN Inc."
__copyright__ = "Copyright 2024, Advanced Systems for Power Engineering Inc."
__license__   = "All rights reserved"
__category__  = "OneLiner"
__pyManager__ = "yes"
__version__   = "1.0.0"
__email__     = "support@aspeninc.com"
__status__    = "Release"

# import OlxAPI
try:
    # running inside OneLiner
    import OlxObj
    runmode = 'internal'
except:
    # running outside of OneLiner
    # OlxAPI path
    olxpath   = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15'
    olxpathpy = 'C:\\Program Files (x86)\\ASPEN\\1LPFv15\\OlxAPI\\python'

    import os,sys
    sys.path.insert(0,olxpath)
    sys.path.insert(0,olxpathpy)
    import OlxObj
    OlxObj.load_olxapi(olxpath)
    runmode = 'external'

import os

import argparse
PARSER_INPUTS = argparse.ArgumentParser(epilog= "")
PARSER_INPUTS.usage =\
    "\n  1. Open OneLiner file (.olr)\
     \n  2. Launch this PrefaultCheck script\
     \n  3. Select path to create output .CSV file when prompted\
     \n  4. Choose whether to open the output .CSV file when prompted"
PARSER_INPUTS.add_argument('-None', help = 'Click Launch Button to run on the current OneLiner .olr file', default = 0, type=int, metavar='')

ARGVS = PARSER_INPUTS.parse_known_args()[0]

# CORE--------------------------------------------------------------------------
def run_prefault_check():
    
    # tools
    import time
    from tkinter import filedialog
    from tkinter import messagebox

    if runmode == 'external':
        oneliner_case = filedialog.askopenfilename(title='Open OneLiner Case (.olr)',filetypes=[("OneLiner", ".olr")])
        OlxObj.OLCase.open(oneliner_case,1) # 1 = read only

    if OlxObj.OLCase.olrFile:
        
        # turn off verbose messaging
        OlxObj.setVerbose(False)
    
        

        # change dir to import checking functions
        original_dir = os.getcwd()
        os.chdir(os.path.dirname(__file__))

        # import checking functions
        from prefault_check.check_ibr_for_nonzero_mw import check_ibr_for_nonzero_mw
        from prefault_check.check_vccs_for_nonzero_prefault_current import check_vccs_for_nonzero_prefault_current
        from prefault_check.check_prefault_voltage_range import check_prefault_voltage_range
        from prefault_check.check_gen_for_nonzero_prefault_current import check_gen_for_nonzero_prefault_current
        from prefault_check.check_gen_for_vref import check_gen_for_vref
        from prefault_check.check_xfmr_for_off_nominal_taps import check_xfmr_for_off_nominal_taps

        # return to original dir
        os.chdir(original_dir)

        # choose output file path
        output_file = filedialog.asksaveasfilename(title='Save CSV file (.csv)',filetypes=[("Comma Separated Values", ".csv")])

        # ensure valid output file path selected
        if not output_file:
            print('Error, invalid output file')
        else:
            if not output_file[-4:].lower() == '.csv':
                output_file += '.csv'

            # anomaly counter
            anomaly_index = 1
            
            # run the checking functions and create csv summary
            # open csv for reporting
            with open(output_file,'w') as f:

                # tool summary
                f.write('ASPEN OneLiner Prefault Voltage Check (Release V%s)\n\n'%(__version__))
                
                f.write('Prefault Voltage Profile Factors Check\n')
                f.write('    This tool aims to identify factors\n')
                f.write('    that may contribute to prefault\n')
                f.write('    voltage profile issues\n\n')

                f.write('Important Note:\n')
                f.write('"    The items reported here are not necessarily"\n')
                f.write('"    modeling errors, but instead represent items"\n')
                f.write('"    that may be contributing to a prefault voltage"\n')
                f.write('"    profile or convergence issue."\n\n')

                f.write('"Check start time: %s"\n'%(time.asctime()))
                f.write('"OneLiner network model: %s"\n\n'%(OlxObj.OLCase.olrFile))

                # write csv header row
                f.write('item,function,object,area,zone,condition,quantity,expected,actual,notes\n')

                # start summary string
                full_summary = ''

                # check_ibr_for_nonzero_mw
                summary,detail = check_ibr_for_nonzero_mw(OlxObj)
                full_summary += '%s\n'%(summary)
                for k,v in detail.items():
                    row_str = '"%s",'%(anomaly_index) + ','.join(['"' + str(x) + '"' for x in v.values()]) + '\n'
                    f.write(row_str)
                    anomaly_index += 1

                # check_vccs_for_nonzero_prefault_current
                summary,detail = check_vccs_for_nonzero_prefault_current(OlxObj, voltage_checkpoints_pu=(0.95,1.0,1.05))
                full_summary += '%s\n'%(summary)
                for k,v in detail.items():
                    row_str = '"%s",'%(anomaly_index) + ','.join(['"' + str(x) + '"' for x in v.values()]) + '\n'
                    f.write(row_str)
                    anomaly_index += 1

                # check_prefault_voltage_range
                summary,detail = check_prefault_voltage_range(OlxObj, checking_range=(0.95,1.05))
                full_summary += '%s\n'%(summary)
                for k,v in detail.items():
                    row_str = '"%s",'%(anomaly_index) + ','.join(['"' + str(x) + '"' for x in v.values()]) + '\n'
                    f.write(row_str)
                    anomaly_index += 1

                # check_gen_for_nonzero_prefault_current
                summary,detail = check_gen_for_nonzero_prefault_current(OlxObj, max_allowable_flow_mva=10.0)
                full_summary += '%s\n'%(summary)
                for k,v in detail.items():
                    row_str = '"%s",'%(anomaly_index) + ','.join(['"' + str(x) + '"' for x in v.values()]) + '\n'
                    f.write(row_str)
                    anomaly_index += 1

                # check_gen_for_vref
                summary,detail = check_gen_for_vref(OlxObj, expected_gen_vref=1.0)
                full_summary += '%s\n'%(summary)
                for k,v in detail.items():
                    row_str = '"%s",'%(anomaly_index) + ','.join(['"' + str(x) + '"' for x in v.values()]) + '\n'
                    f.write(row_str)
                    anomaly_index += 1

                # check_xfmr_for_off_nominal_taps
                summary,detail = check_xfmr_for_off_nominal_taps(OlxObj, expected_tap_min_max=(1.0,1.0))
                full_summary += '%s\n'%(summary)
                for k,v in detail.items():
                    row_str = '"%s",'%(anomaly_index) + ','.join(['"' + str(x) + '"' for x in v.values()]) + '\n'
                    f.write(row_str)
                    anomaly_index += 1
                    
            # print summary
            print('\nASPEN OneLiner Prefault Voltage Check (Release V%s)'%(__version__))

            print('\nImportant Note:')
            print('    The items reported here are not necessarily')
            print('    modeling errors, but instead represent items')
            print('    that may be contributing to a prefault voltage')
            print('    profile or convergence issue.')

            print('\nSummary')
            print('-'*50)
            print(full_summary)
            print('Report saved to:')
            print('%s\n\n'%(output_file))

            # ask whether to open output file
            user_wants_to_open_output = messagebox.askquestion('Output Ready', 'Open %s with default Windows program?'%(output_file))
            if user_wants_to_open_output == 'yes':
                os.startfile(output_file)

        # turn on verbose messaging
        OlxObj.setVerbose(True)

    else:
        
        print('You must open a OneLiner file before running the PrefaultCheck script.')

    if runmode == 'external':
        # close OneLiner case
        OlxObj.OLCase.close()

def main():
    run_prefault_check()
    
if __name__ == '__main__':
    main()


