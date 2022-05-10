from win32com.client import Dispatch
from datetime import date

remitdate = date.today().strftime('%m-%d-%Y')

remitFile = r'M:\Capital Markets\Investor Loan Analysis\Ritz\Investor Position\Cashflow Tie Outs\CMLTI ' \
            fr'2021-INV1\2022\Excel Remit\cmlti-2021-inv1-ta-investor-report-{remitdate}.xls '

cashflowTieout = r'M:\Capital Markets\Investor Loan Analysis\Ritz\Investor Position\Cashflow Tie Outs\CMLTI ' \
                 r'2021-INV1\CMLTI 2021-INV1 Actual Cashflow Tieout - Amol.xlsx '

xl = Dispatch("Excel.Application")
xl.Visible = 1
wb1 = xl.Workbooks.Open(remitFile)
wb2 = xl.Workbooks.Open(cashflowTieout)

# fill down row in cashflow tieout file
maxrow = wb2.Worksheets(1).UsedRange.Rows.Count + 1
xl.Range(f"{maxrow}:{maxrow}").Select()
xl.Selection.FillDown()

Master_Dictionary = {13:{"F27":f"C{maxrow}","F28":f"D{maxrow}","F29":f"D{maxrow+1}","F15":f"F{maxrow}"}, #Remit Summary
                     12:{"K10":f"H{maxrow}","K11":f"I{maxrow}","K12":f"J{maxrow}","K13":f"K{maxrow}",
                         "K14":f"L{maxrow}","E28":f"M{maxrow}"}, #Fees and Interest
                     11:{"B35":f"AO{maxrow}","C35":f"AP{maxrow}","F35":f"AR{maxrow}","K35":f"AS{maxrow}", #B1W Prin
                    "B36":f"AU{maxrow}","C36":f"AV{maxrow}","F36":f"AX{maxrow}","K36":f"AY{maxrow}", #B2W Prin
                    "B37":f"BA{maxrow}","C37":f"BB{maxrow}","F37":f"BD{maxrow}","K37":f"BE{maxrow}"}, #B3W Prin
                     10:{"B19":f"BG{maxrow}","C19":f"BH{maxrow}","F19":f"BJ{maxrow}","K19":f"BK{maxrow}", #B4 Prin
                    "B20":f"BM{maxrow}","C20":f"BN{maxrow}","F20":f"BP{maxrow}","K20":f"BQ{maxrow}", #B5 Prin
                    "B21":f"BS{maxrow}","C21":f"BT{maxrow}","F21":f"BV{maxrow}","K21":f"BW{maxrow}"}, #B6 Prin
                     9:{"H35":f"AQ{maxrow}",#B1W Int
                    "H36":f"AW{maxrow}", #B2W Int
                    "H37":f"BC{maxrow}"}, #B3W Int
                     8:{"H19":f"BI{maxrow}",#B1W Int
                    "H20":f"BO{maxrow}", #B2W Int
                    "H21":f"BU{maxrow}"}, #B3W Int
                     16:{"J28":f"DS{maxrow}","L28":f"DS{maxrow+1}","M28":f"DS{maxrow+2}"}} #D60+

xl.Range(f"A{maxrow}").Select()
xl.Selection.Value = str(date.today())

for page, nest in Master_Dictionary.items():
    for key in nest:
        wb1.activate
        wb1.Worksheets(page).Select()
        xl.Range(key).Select()
        xl.Selection.Copy()
        wb2.activate
        xl.Range(nest[key]).PasteSpecial(Paste=-4163)

maxrow = wb2.Worksheets(1).UsedRange.Rows.Count

xl.Range(f"D{maxrow-1}").Select()
xl.Selection.Copy()
xl.Range(f"D{maxrow-2}").PasteSpecial(Paste=-4163,Operation=2)

xl.Range(f"DS{maxrow}").Select()
xl.Selection.Copy()
xl.Range(f"DS{maxrow-2}").PasteSpecial(Paste=-4163,Operation=2)
xl.Range(f"DS{maxrow-1}").Select()
xl.Selection.Copy()
xl.Range(f"DS{maxrow-2}").PasteSpecial(Paste=-4163,Operation=2)

xl.Range(f"DS{maxrow}").Select()
xl.Selection.Value = 100
xl.Selection.Copy()
xl.Range(f"DS{maxrow-2}").PasteSpecial(Paste=-4163,Operation=5)

xl.Range(f"{maxrow-1}:{maxrow}").Select()
xl.Selection.Clear()
xl.Range(f"B{maxrow}").Select()

xl.Range(f"{maxrow-1}:{maxrow}").Select()
xl.Selection.EntireRow.Delete()