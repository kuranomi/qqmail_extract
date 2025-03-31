from numpy.core.numeric import full
from download_process import QQEmail
from extract_process import get_data, inset_image_and_export, is_gfile, image_extract
from config.settings import EXTRACT_DIR
import pandas as pd

mail = QQEmail()
target_ids = mail.fetch_mails()
for eid in target_ids[3:]:
    print("=" * 50)
    print(f"start  with eid {eid}")
    mail_content = mail.get_content(eid)
    if not mail_content.get("attachemnt_path"):
        print("no attachment in email")
        continue

    df_list = []
    print("handle attachments.....")
    for att_path in mail_content["attachemnt_path"]:
        with pd.ExcelFile(att_path) as xls:
            sheet_names = xls.sheet_names
            print(f"contain sheets: {sheet_names}")
        for sheet_name in sheet_names:
            clean_table = get_data(xls, sheet_name)
            df_list.append(clean_table)
    if len(df_list) > 1:
        fulltable = pd.concat(df_list, ignore_index=True)
        fulltable.reset_index(inplace=True, drop=True)
    else:
        fulltable = df_list[0]

    for c in ["discount", "release_date", "cutoff_date"]:
        if mail_content.get(c) and fulltable[c].isna().all():
            fulltable[c] = mail_content.get(c)
    
    save_dir = mail_content["save_dir"]
    if mail_content.get("image_links"):
        all_path = image_extract(mail_content["image_links"], save_dir)
        if all_path:
            inset_image_and_export(fulltable, save_dir, all_path)
        else:
            fulltable.to_excel(f'{save_dir}/products.xlsx', index=True)
    else:
        fulltable.to_excel(f'{save_dir}/products.xlsx', index=True)
    
    del fulltable
    df_list = []
    break


