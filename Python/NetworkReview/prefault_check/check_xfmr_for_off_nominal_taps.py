def check_xfmr_for_off_nominal_taps(OlxObj,
                       expected_tap_min_max=(1.0,1.0)):
    
    """
    Check for transformers with non-unity tap settings

    Inputs:
        OlxObj:
            instance of the ASPEN OneLiner OlxObj API
            where the OLCase object is already loaded

        expected_tap_min_max:
            tuple of min,max allowable tap ratio settings

    Outputs:
        summary: A printable string containing basic summary information
                 about the checking results
        detail_dict: A dictionary containing the checking results details
    """

    detail_dict = {}
    detail_index = 0

    # check for off-nominal transformer taps
    for xfmr in OlxObj.OLCase.XFMR:
        b1_nominal_kv = xfmr.BUS1.KV
        b2_nominal_kv = xfmr.BUS2.KV
        b1_tap = xfmr.PRITAP
        b2_tap = xfmr.SECTAP

        if xfmr.CONFIGP in ['D','E'] and xfmr.CONFIGS in ['D','E']:
             b1_nominal_kv = b1_nominal_kv * (3**0.5) # delta delta requires sqrt(3)
             b2_nominal_kv = b2_nominal_kv * (3**0.5) # delta delta requires sqrt(3)
        
        pri_ratio = round(b1_tap / b1_nominal_kv, 2)
        sec_ratio = round(b2_tap / b2_nominal_kv, 2)

        if pri_ratio < min(expected_tap_min_max):
            detail_dict[detail_index] = {'function':'check_xfmr_for_off_nominal_taps(expected_tap_min_max=%s)'%(str(expected_tap_min_max)),
                                            'object':xfmr.toString(),
                                            'area':xfmr.BUS1.AREANO,
                                            'zone':xfmr.BUS1.ZONENO,
                                            'condition':'xfmr_pri_tap_low',
                                            'quantity':'xfmr.pri.tap',
                                            'expected':min(expected_tap_min_max),
                                            'actual':pri_ratio,
                                            'notes':'Review tap settings in the context of prefault voltage profile'}
            detail_index += 1

        if pri_ratio > max(expected_tap_min_max):
            detail_dict[detail_index] = {'function':'check_xfmr_for_off_nominal_taps(expected_tap_min_max=%s)'%(str(expected_tap_min_max)),
                                            'object':xfmr.toString(),
                                            'area':xfmr.BUS1.AREANO,
                                            'zone':xfmr.BUS1.ZONENO,
                                            'condition':'xfmr_pri_tap_high',
                                            'quantity':'xfmr.pri.tap',
                                            'expected':max(expected_tap_min_max),
                                            'actual':pri_ratio,
                                            'notes':'Review tap settings in the context of prefault voltage profile'}
            detail_index += 1

        if sec_ratio < min(expected_tap_min_max):
            detail_dict[detail_index] = {'function':'check_xfmr_for_off_nominal_taps(expected_tap_min_max=%s)'%(str(expected_tap_min_max)),
                                            'object':xfmr.toString(),
                                            'area':xfmr.BUS2.AREANO,
                                            'zone':xfmr.BUS2.ZONENO,
                                            'condition':'xfmr_sec_tap_low',
                                            'quantity':'xfmr.sec.tap',
                                            'expected':min(expected_tap_min_max),
                                            'actual':sec_ratio,
                                            'notes':'Review tap settings in the context of prefault voltage profile'}
            detail_index += 1
        if sec_ratio > max(expected_tap_min_max):
            detail_dict[detail_index] = {'function':'check_xfmr_for_off_nominal_taps(expected_tap_min_max=%s)'%(str(expected_tap_min_max)),
                                            'object':xfmr.toString(),
                                            'area':xfmr.BUS2.AREANO,
                                            'zone':xfmr.BUS2.ZONENO,
                                            'condition':'xfmr_sec_tap_high',
                                            'quantity':'xfmr.sec.tap',
                                            'expected':max(expected_tap_min_max),
                                            'actual':sec_ratio,
                                            'notes':'Review tap settings in the context of prefault voltage profile'}
            detail_index += 1
    
    # check for off-nominal transformer taps 3-winding
    for xfmr in OlxObj.OLCase.XFMR3:
        b1_nominal_kv = xfmr.BUS1.KV
        b2_nominal_kv = xfmr.BUS2.KV
        b3_nominal_kv = xfmr.BUS3.KV
        b1_tap = xfmr.PRITAP
        b2_tap = xfmr.SECTAP
        b3_tap = xfmr.TERTAP

        if xfmr.CONFIGP in ['D','E'] and xfmr.CONFIGS in ['D','E'] and xfmr.CONFIGT in ['D','E']:
             b1_nominal_kv = b1_nominal_kv * (3**0.5) # delta delta requires sqrt(3)
             b2_nominal_kv = b2_nominal_kv * (3**0.5) # delta delta requires sqrt(3)
             b3_nominal_kv = b3_nominal_kv * (3**0.5) # delta delta requires sqrt(3)
        
        pri_ratio = round(b1_tap / b1_nominal_kv, 2)
        sec_ratio = round(b2_tap / b2_nominal_kv, 2)
        ter_ratio = round(b3_tap / b3_nominal_kv, 2)

        if pri_ratio < min(expected_tap_min_max):
            detail_dict[detail_index] = {'function':'check_xfmr_for_off_nominal_taps(expected_tap_min_max=%s)'%(str(expected_tap_min_max)),
                                            'object':xfmr.toString(),
                                            'area':xfmr.BUS1.AREANO,
                                            'zone':xfmr.BUS1.ZONENO,
                                            'condition':'xfmr_pri_tap_low',
                                            'quantity':'xfmr.pri.tap',
                                            'expected':min(expected_tap_min_max),
                                            'actual':pri_ratio,
                                            'notes':'Review tap settings in the context of prefault voltage profile'}
            detail_index += 1

        if pri_ratio > max(expected_tap_min_max):
            detail_dict[detail_index] = {'function':'check_xfmr_for_off_nominal_taps(expected_tap_min_max=%s)'%(str(expected_tap_min_max)),
                                            'object':xfmr.toString(),
                                            'area':xfmr.BUS1.AREANO,
                                            'zone':xfmr.BUS1.ZONENO,
                                            'condition':'xfmr_pri_tap_high',
                                            'quantity':'xfmr.pri.tap',
                                            'expected':max(expected_tap_min_max),
                                            'actual':pri_ratio,
                                            'notes':'Review tap settings in the context of prefault voltage profile'}
            detail_index += 1

        if sec_ratio < min(expected_tap_min_max):
            detail_dict[detail_index] = {'function':'check_xfmr_for_off_nominal_taps(expected_tap_min_max=%s)'%(str(expected_tap_min_max)),
                                            'object':xfmr.toString(),
                                            'area':xfmr.BUS2.AREANO,
                                            'zone':xfmr.BUS2.ZONENO,
                                            'condition':'xfmr_sec_tap_low',
                                            'quantity':'xfmr.sec.tap',
                                            'expected':min(expected_tap_min_max),
                                            'actual':sec_ratio,
                                            'notes':'Review tap settings in the context of prefault voltage profile'}
            detail_index += 1
        if sec_ratio > max(expected_tap_min_max):
            detail_dict[detail_index] = {'function':'check_xfmr_for_off_nominal_taps(expected_tap_min_max=%s)'%(str(expected_tap_min_max)),
                                            'object':xfmr.toString(),
                                            'area':xfmr.BUS2.AREANO,
                                            'zone':xfmr.BUS2.ZONENO,
                                            'condition':'xfmr_sec_tap_high',
                                            'quantity':'xfmr.sec.tap',
                                            'expected':max(expected_tap_min_max),
                                            'actual':sec_ratio,
                                            'notes':'Review tap settings in the context of prefault voltage profile'}
            detail_index += 1
        
        if ter_ratio < min(expected_tap_min_max):
            detail_dict[detail_index] = {'function':'check_xfmr_for_off_nominal_taps(expected_tap_min_max=%s)'%(str(expected_tap_min_max)),
                                            'object':xfmr.toString(),
                                            'area':xfmr.BUS3.AREANO,
                                            'zone':xfmr.BUS3.ZONENO,
                                            'condition':'xfmr_ter_tap_low',
                                            'quantity':'xfmr.ter.tap',
                                            'expected':min(expected_tap_min_max),
                                            'actual':ter_ratio,
                                            'notes':'Review tap settings in the context of prefault voltage profile'}
            detail_index += 1
        if ter_ratio > max(expected_tap_min_max):
            detail_dict[detail_index] = {'function':'check_xfmr_for_off_nominal_taps(expected_tap_min_max=%s)'%(str(expected_tap_min_max)),
                                            'object':xfmr.toString(),
                                            'area':xfmr.BUS3.AREANO,
                                            'zone':xfmr.BUS3.ZONENO,
                                            'condition':'xfmr_ter_tap_high',
                                            'quantity':'xfmr.ter.tap',
                                            'expected':max(expected_tap_min_max),
                                            'actual':ter_ratio,
                                            'notes':'Review tap settings in the context of prefault voltage profile'}
            detail_index += 1
            
    summary = 'Found %s transformer(s) with taps outside of range %s'%(str(len(detail_dict)),str(expected_tap_min_max))

    return summary,detail_dict

def main():
    pass

