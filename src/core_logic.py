import pandas as pd
import numpy as np
from src.config import BOSTON_ZONES, EXCLUDED_PLANTS

def excelExtract(file, sheet):
    """Extracts Flows and Dispatch data from the provided excel sheet."""
    try:
        excelContent = pd.read_excel(file, sheet_name=sheet)
        
        # Locate the specific rows that delimit the data blocks
        flows_indices = excelContent.index[excelContent.iloc[:,0] == 'Flows'].tolist()
        dispatch_indices = excelContent.index[excelContent.iloc[:,0] == 'Dispatch'].tolist()
        
        if not flows_indices or not dispatch_indices:
            raise ValueError(f"Sheet '{sheet}' is missing 'Flows' or 'Dispatch' headers in the first column.")
            
        startingIndex_Flows = flows_indices[0] + 1
        startingIndex_Dispatch = dispatch_indices[0] + 1
        endingIndex_Flows = startingIndex_Dispatch - 3
        
        flows = pd.DataFrame(excelContent.iloc[startingIndex_Flows+1:endingIndex_Flows, 0:14])
        flows.columns = excelContent.iloc[startingIndex_Flows, 0:14]
        
        dispatch = pd.DataFrame(excelContent.iloc[startingIndex_Dispatch+1:, :])
        dispatch.columns = excelContent.iloc[startingIndex_Dispatch, :]
        dispatch = dispatch.reset_index(drop=True)
        
        return flows, dispatch
        
    except Exception as e:
        raise Exception(f"Failed to extract data from {file} - {sheet}: {str(e)}")

def tableReformation(data, flowData, busNo, projectSize):
    """Calculates the dependent parameters after changing the values in the 'Pnew' column."""
    try:
        data = data.copy()
        flowData = flowData.copy()
        
        data.loc[data['  Bus#'] == busNo, 'Pnew'] = projectSize
        
        pDelta = data.loc[data['  Bus#'] == busNo, 'Pnew'] - data.loc[data['  Bus#'] == busNo, 'Pgen or Pload']
        data.loc[data['  Bus#'] == busNo, 'Pdelta'] = pDelta
        
        pAvail = data.loc[data['  Bus#'] == busNo, 'PMax Gen'] - data.loc[data['  Bus#'] == busNo, 'Pnew']
        data.loc[data['  Bus#'] == busNo, 'Pavail'] = pAvail
        
        for j in range(1, flowData.shape[0] + 1):
            temp = f'Impact_{j}'
            impactVal = pDelta * data.loc[data['  Bus#'] == busNo, f'Dfax_{j}']
            data.loc[data['  Bus#'] == busNo, temp] = impactVal
            
            # calculate flowChange
            flowChange = data[temp].sum()
            flowData.loc[j+2, 'FlowChange'] = flowChange
            
            # calculate flowRes
            flowRes = flowData.loc[j+2, 'FlowChange'] + flowData.loc[j+2, 'FlowInit']
            flowData.loc[j+2, 'FlowRes'] = flowRes
            
            # calculate loading
            flowData.loc[j+2, 'Loading'] = flowData.loc[j+2, 'FlowRes'] / flowData.loc[j+2, 'Limit']
            
        return flowData, data
    except Exception as e:
        raise Exception(f"Failed to reform table: {str(e)}")

def is_excluded(bus_name):
    if not isinstance(bus_name, str):
        return False
    return any(excluded in bus_name for excluded in EXCLUDED_PLANTS)

def redispatch(flows, dispatch, busNo, charging, switchNo=0):
    """Performs the redispatch logic based on capacities and constraints."""
    flows = flows.copy()
    dispatch = dispatch.copy()
    
    for i in range(dispatch.shape[0]):
        pNew = dispatch.iloc[i].loc['Pnew']
        pmaxGen = dispatch.iloc[i].loc['PMax Gen']
        pavail = dispatch.iloc[i].loc['Pavail']
        busName = dispatch.iloc[i].loc['Bus Name    ']
        pDeltaSum = dispatch['Pdelta'].sum()
        
        droppingCapacity = 0
        
        if charging == 'N':
            if pDeltaSum <= 0:
                break
            if pNew == 0:
                continue
            droppingCapacity = pNew - pDeltaSum if pNew > pDeltaSum else 0
            
        elif charging == 'Y':
            if pDeltaSum <= 0:
                if pavail != 0:
                    if pmaxGen > pavail:
                        droppingCapacity = pNew + pavail if abs(pDeltaSum) > pavail else pNew + abs(pDeltaSum)
                    else:
                        droppingCapacity = pNew + pmaxGen if abs(pDeltaSum) > pmaxGen else pNew + abs(pDeltaSum)
                else:
                    continue
            else:
                break
                
        busN = dispatch.iloc[i].loc['  Bus#']
        zoneN_series = dispatch.loc[dispatch['  Bus#'] == busN, 'Zone']
        zoneN = int(zoneN_series.iloc[0]) if not zoneN_series.empty else None
        
        if busN != busNo and not is_excluded(busName):
            if switchNo != 0:
                if zoneN in BOSTON_ZONES:
                    flows, dispatch = tableReformation(dispatch, flows, busN, droppingCapacity)
            else:
                flows, dispatch = tableReformation(dispatch, flows, busN, droppingCapacity)

    # Second pass for remaining delta
    pDeltaSum2 = dispatch['Pdelta'].sum()
    if pDeltaSum2 > 0:
        for i in range(dispatch.shape[0]):
            pDeltaSum2 = dispatch['Pdelta'].sum()
            if pDeltaSum2 <= 0:
                break
                
            busName = dispatch.iloc[i].loc['Bus Name    ']
            pNew = dispatch.iloc[i].loc['Pnew']
            
            if pNew == 0:
                continue
                
            droppingCapacity = pNew - pDeltaSum2 if pNew > pDeltaSum2 else 0
            
            if switchNo != 0:
                busN = dispatch.iloc[i].loc['  Bus#']
                zoneN_series = dispatch.loc[dispatch['  Bus#'] == busN, 'Zone']
                zoneN = int(zoneN_series.iloc[0]) if not zoneN_series.empty else None
                
                dfaxN_series = dispatch.loc[dispatch['  Bus#'] == busN, 'Dfax_1']
                dfaxN = float(dfaxN_series.iloc[0]) if not dfaxN_series.empty else 0.0
                
                if zoneN not in BOSTON_ZONES and (-0.01 <= dfaxN <= 0.01):
                    if busN != busNo and not is_excluded(busName):
                        flows, dispatch = tableReformation(dispatch, flows, busN, droppingCapacity)

    return flows, dispatch
