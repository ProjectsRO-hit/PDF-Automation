import pandas as pd
from fillpdf import fillpdfs
import os

df = pd.read_excel('form_fields.xlsx')

input_pdf_path = 'Source.pdf'
output_directory = 'output_pdfs'

# Create the output directory if it doesn't exist
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

form_fields = fillpdfs.get_form_fields(input_pdf_path)
print("PDF Form Fields: ", form_fields)

for index, row in df.iterrows():
    # Map the actual form fields to your data from Excel
    data_dict = {
        # 'text_1gfsf': row['Agency'],   
        'text_1gfsf': row['Name'] if 'Name' in df.columns else '',       
        'text_2zeqg':row['Mailing Address'] if 'Mailing Address' in df.columns else '',       
        'text_3soga':row['Garaging Address'] if 'Garaging Address' in df.columns else '',            
        'text_4tciq':row['DBA'] if 'DBA' in df.columns else '', 
        'text_5lbbe':row['Contact Name'] if 'Contact Name' in df.columns else '',           
        'text_6vpwl':row['email'] if 'email' in df.columns else '', 
        'text_7puhe':row['Target Effective'] if 'Target Effective' in df.columns else '', 
        'text_8faqb':row['Phone Number'] if 'Phone Number' in df.columns else '', 
        'text_9mnc':row['SMS Pin'] if 'SMS Pin' in df.columns else '', 
        'text_10bgki': row['Owners/Executives'] if 'Owners/Executives' in df.columns else '', 
        'text_11wder': row['Owners/Executives1'] if 'Owners/Executives1' in df.columns else '', 
        'text_12wcod': row['Owners/Executives2'] if 'Owners/Executives2' in df.columns else '', 
        'text_13txco': row['Target Premium'] if 'Target Premium' in df.columns else '', 
        'text_14odgw': row['Federal Tax ID Number'] if 'Federal Tax ID Number' in df.columns else '', 
        'text_15qcfx': row['UDT1'] if 'UDT1' in df.columns else '', 
        'text_16yccq': row['MC Number'] if 'MC Number' in df.columns else '', 
        'text_20qxeb': row['Number of Years in Business'] if 'Number of Years in Business' in df.columns else '', 
        'text_27ymlc': row['UDT2'] if 'UDT2' in df.columns else '', 
        'text_28hx': row['MC Number2'] if 'MC Number2' in df.columns else '', 
        'text_100bafa': row['Agc1'] if 'Agc1' in df.columns else '',
        'text_101itmu': row['Prod1'] if 'Prod1' in df.columns else '', 
        'text_31fnne': row['F22'] if 'F22' in df.columns else '', 
        'text_34nfsr': row['F23'] if 'F23' in df.columns else '', 
        'text_35syw': row['F24'] if 'F24' in df.columns else '', 
        'text_36lxxa': row['F25'] if 'F25' in df.columns else '', 
        'text_37dadg': row['F26'] if 'F26' in df.columns else '', 
        'text_38lviu': row['F27'] if 'F27' in df.columns else '', 
        'text_39pkoq': row['TOE28'] if 'TOE28' in df.columns else '', 
        'text_40idty': row['TOE29'] if 'TOE29' in df.columns else '', 
        'text_41oitp': row['TOE30'] if 'TOE30' in df.columns else '', 
        'text_43xzcb': row['TOE31'] if 'TOE31' in df.columns else '', 
        'text_44tuy': row['TOE32'] if 'TOE32' in df.columns else '', 
        'text_45bixs': row['VIN33'] if 'VIN33' in df.columns else '', 
        'text_46vxwy': row['VIN34'] if 'VIN34' in df.columns else '', 
        'text_47fq': row['VIN35'] if 'VIN35' in df.columns else '', 
        'text_48bppp': row['VIN36'] if 'VIN36' in df.columns else '', 
        'text_49hsyl': row['VIN37'] if 'VIN37' in df.columns else '', 
        'text_50lkqu': row['Y38'] if 'Y38' in df.columns else '', 
        'text_51bbvs': row['Y39'] if 'Y39' in df.columns else '', 
        'text_52ohui': row['Y40'] if 'Y40' in df.columns else '', 
        'text_53jxhf': row['Y41'] if 'Y41' in df.columns else '', 
        'text_54gqgh': row['Y42'] if 'Y42' in df.columns else '', 
        'text_55awga': row['M43'] if 'M43' in df.columns else '', 
        'text_56exud': row['M44'] if 'M44' in df.columns else '', 
        'text_57lrei': row['M45'] if 'M45' in df.columns else '', 
        'text_58lyd': row['M46'] if 'M46' in df.columns else '', 
        'text_59twur': row['M47'] if 'M47' in df.columns else '', 
        'text_60fvwb': row['O48'] if 'O48' in df.columns else '', 
        'text_61xkjv': row['O49'] if 'O49' in df.columns else '', 
        'text_62rmfz': row['O50'] if 'O50' in df.columns else '', 
        'text_63fbnf': row['O51'] if 'O51' in df.columns else '', 
        'text_64ex': row['O52'] if 'O52' in df.columns else '', 
        'text_65bsyq': row['DN53'] if 'DN53' in df.columns else '', 
        'text_66prns': row['DN54'] if 'DN54' in df.columns else '', 
        'text_67knfh': row['DN55'] if 'DN55' in df.columns else '', 
        'text_68xjwp': row['DN56'] if 'DN56' in df.columns else '', 
        'text_69hijl': row['DN57'] if 'DN57' in df.columns else '', 
        'text_70tsmr': row['DN58'] if 'DN58' in df.columns else '', 
        'text_71fwra': row['DN59'] if 'DN59' in df.columns else '', 
        'text_72nbys': row['DN60'] if 'DN60' in df.columns else '', 
        'text_73ramk': row['DOB61'] if 'DOB61' in df.columns else '', 
        'text_74az': row['DOB62'] if 'DOB62' in df.columns else '', 
        'text_75prju': row['DOB63'] if 'DOB63' in df.columns else '', 
        'text_76wjam': row['DOB64'] if 'DOB64' in df.columns else '', 
        'text_77cbsd': row['DOB65'] if 'DOB65' in df.columns else '', 
        'text_78rwkj': row['DOB66'] if 'DOB66' in df.columns else '', 
        'text_79eqzv': row['DOB67'] if 'DOB67' in df.columns else '', 
        'text_80jbxb': row['DOB68'] if 'DOB68' in df.columns else '', 
        'text_81wlqq': row['DL69'] if 'DL69' in df.columns else '', 
        'text_82vsvq': row['DL70'] if 'DL70' in df.columns else '', 
        'text_83pili': row['DL71'] if 'DL71' in df.columns else '', 
        'text_84lold': row['DL72'] if 'DL72' in df.columns else '', 
        'text_85osug': row['DL73'] if 'DL73' in df.columns else '', 
        'text_86adxf': row['DL74'] if 'DL74' in df.columns else '', 
        'text_87pteg': row['DL75'] if 'DL75' in df.columns else '', 
        'text_88qggz': row['DL76'] if 'DL76' in df.columns else '', 
        'text_89rati': row['YOE77'] if 'YOE77' in df.columns else '', 
        'text_90diqe': row['YOE78'] if 'YOE78' in df.columns else '', 
        'text_91ossm': row['YOE79'] if 'YOE79' in df.columns else '', 
        'text_92pvrb': row['YOE80'] if 'YOE80' in df.columns else '', 
        'text_93vpy':row['YOE81'] if 'YOE81' in df.columns else '', 
        'text_94fkmc': row['YOE82'] if 'YOE82' in df.columns else '',
        'text_95xyvj': row['YOE83'] if 'YOE83' in df.columns else '', 
        'text_96exok': row['YOE84'] if 'YOE84' in df.columns else '', 
        # 'checkbox_112loff': 'Yes_ajdq', 
        'text_97efjn': row['AppN86'] if 'AppN86' in df.columns else '', 
        'text_98wbpf': row['Titl87'] if 'Titl87' in df.columns else '', 
        'text_99musr': row['Date88'] if 'Date88' in df.columns else ''
    }

    output_pdf_path = os.path.join(output_directory, f'filled_form_{index + 1}.pdf')

    fillpdfs.write_fillable_pdf(input_pdf_path, output_pdf_path, data_dict)

    print(f'Generated {output_pdf_path}')

print("All PDFs generated successfully!")

