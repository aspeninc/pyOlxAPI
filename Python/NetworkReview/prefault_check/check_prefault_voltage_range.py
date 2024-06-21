def check_prefault_voltage_range(OlxObj,
                                   checking_range=(0.95,1.05)):
    
    """
    Check models for prefault voltages that are outside
    of the defined range

    Inputs:
        OlxObj:
            instance of the ASPEN OneLiner OlxObj API
            where the OLCase object is already loaded

        checking_range:
            tuple of min,max per unit voltage ranges allowable
            at each object terminal bus

    Outputs:
        summary: A printable string containing basic summary information
                 about the checking results
        detail_dict: A dictionary containing the checking results details
    """

    detail_dict = {}
    detail_index = 0

    # find high kV fault bus
    max_kv = 0
    for bus in OlxObj.OLCase.BUS:
        if bus.KV > max_kv:
            max_kv = bus.KV
            fault_bus = bus
            
    # define fault options
    fault_spec = OlxObj.SPEC_FLT.Classical(obj=fault_bus,fltApp='BUS',fltConn='3LG',Z=[9999.0,9999.0])

    # simulate fault with defined options
    OlxObj.OLCase.simulateFault(fault_spec,1) # (0/1) 1 = clear previous result flag

    term1_types = ['GEN', 'LOAD', 'SHUNT', 'GENW3', 'GENW4', 'CCGEN', 'SVD']
    for fault_simulation_result in OlxObj.FltSimResult:
        for gen in OlxObj.OLCase.getData(term1_types):
            ibr_bus_voltage_pu = abs(fault_simulation_result.voltageSeq(gen.BUS)[1])/(gen.BUS.KV/(3.0**(0.5)))
            if ibr_bus_voltage_pu < min(checking_range) or ibr_bus_voltage_pu > max(checking_range):
                if ibr_bus_voltage_pu < min(checking_range):
                    if ibr_bus_voltage_pu == 0:
                        pass # these must be island buses
                    else:
                        detail_dict[detail_index] = {'function':'check_prefault_voltage_range(checking_range=%s)'%(str(checking_range)),
                                           'object':gen.toString(),
                                           'area':gen.BUS.AREANO,
                                           'zone':gen.BUS.ZONENO,
                                           'condition':'low_voltage',
                                           'quantity':'bus.voltage.pu',
                                           'expected':'>=%s'%(min(checking_range)),
                                           'actual':ibr_bus_voltage_pu,
                                           'notes':'Simulation converged: %s'%(fault_simulation_result.CONVERGED)}
                        detail_index += 1

                if ibr_bus_voltage_pu > max(checking_range):

                    detail_dict[detail_index] = {'function':'check_prefault_voltage_range(checking_range=%s)'%(str(checking_range)),
                                           'object':gen.toString(),
                                           'area':gen.BUS.AREANO,
                                           'zone':gen.BUS.ZONENO,
                                           'condition':'high_voltage',
                                           'quantity':'bus.voltage.pu',
                                           'expected':'<=%s'%(max(checking_range)),
                                           'actual':ibr_bus_voltage_pu,
                                           'notes':'Simulation converged: %s'%(fault_simulation_result.CONVERGED)}
                    detail_index += 1
    
    term2_types = ['SERIESRC','SHIFTER']
    for fault_simulation_result in OlxObj.FltSimResult:
        for gen in OlxObj.OLCase.getData(term2_types):
            for bus in gen.bus:
                ibr_bus_voltage_pu = abs(fault_simulation_result.voltageSeq(bus)[1])/(bus.KV/(3.0**(0.5)))
                if ibr_bus_voltage_pu < min(checking_range) or ibr_bus_voltage_pu > max(checking_range):
                    if ibr_bus_voltage_pu < min(checking_range):
                        if ibr_bus_voltage_pu == 0:
                            pass # these must be island buses
                        else:
                            detail_dict[detail_index] = {'function':'check_prefault_voltage_range(checking_range=%s)'%(str(checking_range)),
                                            'object':'%s @ BUS: %s %skV'%(gen.toString(),bus.NAME,round(bus.KV,4)),
                                            'area':bus.AREANO,
                                            'zone':bus.ZONENO,
                                            'condition':'low_voltage',
                                            'quantity':'bus.voltage.pu',
                                            'expected':'>=%s'%(min(checking_range)),
                                            'actual':ibr_bus_voltage_pu,
                                            'notes':'Simulation converged: %s'%(fault_simulation_result.CONVERGED)}
                            detail_index += 1

                    if ibr_bus_voltage_pu > max(checking_range):

                        detail_dict[detail_index] = {'function':'check_prefault_voltage_range(checking_range=%s)'%(str(checking_range)),
                                            'object':'%s @ BUS: %s %skV'%(gen.toString(),bus.NAME,round(bus.KV,4)),
                                            'area':bus.AREANO,
                                            'zone':bus.ZONENO,
                                            'condition':'high_voltage',
                                            'quantity':'bus.voltage.pu',
                                            'expected':'<=%s'%(max(checking_range)),
                                            'actual':ibr_bus_voltage_pu,
                                            'notes':'Simulation converged: %s'%(fault_simulation_result.CONVERGED)}
                        detail_index += 1
    
    term3_types = ['XFMR','XFMR3']
    for fault_simulation_result in OlxObj.FltSimResult:
        for gen in OlxObj.OLCase.getData(term3_types):
            for bus in gen.bus:
                ibr_bus_voltage_pu = abs(fault_simulation_result.voltageSeq(bus)[1])/(bus.KV/(3.0**(0.5)))
                if ibr_bus_voltage_pu < min(checking_range) or ibr_bus_voltage_pu > max(checking_range):
                    if ibr_bus_voltage_pu < min(checking_range):
                        if ibr_bus_voltage_pu == 0:
                            pass # these must be island buses
                        else:
                            detail_dict[detail_index] = {'function':'check_prefault_voltage_range(checking_range=%s)'%(str(checking_range)),
                                            'object':'%s @ BUS: %s %skV'%(gen.toString(),bus.NAME,round(bus.KV,4)),
                                            'area':bus.AREANO,
                                            'zone':bus.ZONENO,
                                            'condition':'low_voltage',
                                            'quantity':'bus.voltage.pu',
                                            'expected':'>=%s'%(min(checking_range)),
                                            'actual':ibr_bus_voltage_pu,
                                            'notes':'Off-nominal tap settings may affect this voltage'}
                            detail_index += 1

                    if ibr_bus_voltage_pu > max(checking_range):

                        detail_dict[detail_index] = {'function':'check_prefault_voltage_range(checking_range=%s)'%(str(checking_range)),
                                            'object':'%s @ BUS: %s %skV'%(gen.toString(),bus.NAME,round(bus.KV,4)),
                                            'area':bus.AREANO,
                                            'zone':bus.ZONENO,
                                            'condition':'high_voltage',
                                            'quantity':'bus.voltage.pu',
                                            'expected':'<=%s'%(max(checking_range)),
                                            'actual':ibr_bus_voltage_pu,
                                            'notes':'Off-nominal tap settings may affect this voltage'}
                        detail_index += 1

    summary = 'Found %s prefault bus voltage(s) outside of range %s pu'%(str(len(detail_dict)),str(checking_range))

    return summary,detail_dict

def main():
    pass

