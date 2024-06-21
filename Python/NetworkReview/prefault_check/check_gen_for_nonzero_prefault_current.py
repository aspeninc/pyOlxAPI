def check_gen_for_nonzero_prefault_current(OlxObj,
                                            max_allowable_flow_mva=10.0):
    
    """
    Check generator models for prefault mva flows

    Inputs:
        OlxObj:
            instance of the ASPEN OneLiner OlxObj API
            where the OLCase object is already loaded

        max_allowable_flow_mva:
            float of maximum allowable prefault generator mva flow

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
    fault_spec = OlxObj.SPEC_FLT.Classical(obj=fault_bus,fltApp='BUS',fltConn='3LG',Z=[999999.0,999999.0])
    # simulate fault with defined options
    OlxObj.OLCase.simulateFault(fault_spec,1) # (0/1) 1 = clear previous result flag

    gen_prefault_flows = {}
    gen_prefault_flows_mva = {}
    for fault_simulation_result in OlxObj.FltSimResult:
        for gen in OlxObj.OLCase.GEN:
            if gen.FLAG == 1 and len(gen.BUS.findBusNeibor(5)) > 5: # online and not islanded
                gen_current_mag = abs(fault_simulation_result.currentSeq(gen)[1])
                gen_voltage_mag = abs(fault_simulation_result.voltageSeq(gen.BUS)[1]) * 1e3
                gen_prefault_flows[gen.toString()] = gen_current_mag
                gen_prefault_flows_mva[gen.toString()] = gen_current_mag * gen_voltage_mag / 1e6 * 3.0

    for gen,flow in sorted(list(gen_prefault_flows_mva.items()),key=lambda x:x[1],reverse=True):
        try:
            gen_obj = OlxObj.OLCase.findGEN(gen)
        except: # error finding gen, try more reliable, less efficient method
            for genx in OlxObj.OLCase.GEN:
                if genx.toString() == gen:
                    gen_obj = genx
        if flow > max_allowable_flow_mva:
            #print(gen,flow)
            detail_dict[detail_index] = {'function':'check_gen_for_nonzero_prefault_current(max_allowable_flow_mva=%s)'%(str(max_allowable_flow_mva)),
                                           'object':gen,
                                           'area':gen_obj.BUS.AREANO,
                                           'zone':gen_obj.BUS.ZONENO,
                                           'condition':'prefault_gen_flow',
                                           'quantity':'gen.flow.mva',
                                           'expected':'<=%s'%(max_allowable_flow_mva),
                                           'actual':flow,
                                           'notes':'It may be helpful to find the reason for this flow'}
            detail_index += 1
            
    summary = 'Found %s generator(s) with prefault flow greater than %s MVA'%(str(len(detail_dict)),max_allowable_flow_mva)

    return summary,detail_dict

def main():
    pass

