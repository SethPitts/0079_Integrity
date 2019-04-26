import pandas as pd
import text_formatter as tf

# get various sections of integrity file from excel

checks_in_template_df = pd.read_excel('Updated_CTN-0079_Integrity_Tracker.xlsx', sheet_name='0079_Checks')
comments_df = pd.read_excel('Updated_CTN-0079_Integrity_Tracker.xlsx', sheet_name='0079_Checks')
req_name_df = pd.read_excel('Updated_CTN-0079_Integrity_Tracker.xlsx', sheet_name='0079_Checks')
file_name_df = pd.read_excel('Updated_CTN-0079_Integrity_Tracker.xlsx', sheet_name='0079_Checks')
temp_var_df = pd.read_excel('Updated_CTN-0079_Integrity_Tracker.xlsx', sheet_name='TEMPVARS')
checks_df = pd.read_excel('Updated_CTN-0079_Integrity_Tracker.xlsx', sheet_name='0079_Checks')

# Get Checks present in template
checks_in_template = checks_in_template_df.FILE_NAME
for req_file_name in checks_in_template:
    with open('REQs\{}.REQ'.format(req_file_name), 'w') as req_file:
        # Write Comments
        filtered_comments_df = comments_df[(comments_df.FILE_NAME == req_file_name)]  # Filter data frame by check
        comments = filtered_comments_df.COMMENTS_FOR_REQ_FILE
        comments = comments.drop_duplicates(keep='first')
        comment_template = "COMMENT  {}\n"
        for comment in comments:
            if type(comment) != float:
                comment = tf.format_by_charater_length_with_keyword(max_length=50, text=comment, keyword="COMMENT ")
                req_file.write(comment + "\n")

        # write new line
        req_file.write("\n")

        # Write REQ-NAME
        filtered_req_name_df = req_name_df[(req_name_df.FILE_NAME == req_file_name)]  # Filter data frame by check
        filtered_req_name_df = filtered_req_name_df.drop_duplicates(subset='FILE_NAME', keep='first')
        req_name_template = "REQ-NAME {}{}\n"
        for idx, req_info in filtered_req_name_df.iterrows():
            req_name = req_info.loc['FILE_NAME']
            padded_spaces = " " * (10 - len(req_name))
            retain_delete = req_info.loc['RETAIN_DELETE']
            req_file.write(req_name_template.format(req_name + padded_spaces, retain_delete))

        # Write FILE-1-NAME
        filtered_file_1_name_df = file_name_df[(file_name_df.FILE_NAME == req_file_name)]  # Filter data frame by check
        filtered_file_1_name_df = filtered_file_1_name_df.drop_duplicates(subset='FILE_1', keep='first')
        file_name_template = "FILE-{}{}{}\n"
        for idx, file_info in filtered_file_1_name_df.iterrows():
            file_num = "1"
            file_name = file_info.loc['FILE_1']
            padded_spaces_1 = " " * (10 - len(file_name))
            padded_spaces_2 = " " * (4 - len(str(file_num)))
            key_fields = " ".join(file_info.loc['KEYFIELDS_1'].split(","))
            req_file.write(
                file_name_template.format(file_num + padded_spaces_2, file_name + padded_spaces_1, key_fields))

        # Write FILE-2-NAME
        filtered_file_2_name_df = file_name_df[(file_name_df.FILE_NAME == req_file_name)]  # Filter data frame by check
        filtered_file_2_name_df = filtered_file_2_name_df.drop_duplicates(subset='FILE_2', keep='first')
        file_name_template = "FILE-{}{}{}\n"
        for idx, file_info in filtered_file_2_name_df.iterrows():
            file_num = "2"
            file_name = file_info.loc['FILE_2']
            padded_spaces_1 = " " * (10 - len(file_name))
            padded_spaces_2 = " " * (4 - len(str(file_num)))
            key_fields = " ".join(file_info.loc['KEYFIELDS_2'].split(","))
            req_file.write(
                file_name_template.format(file_num + padded_spaces_2, file_name + padded_spaces_1, key_fields))

        # New Line
        req_file.write('\n')

        # Write TEMPFILES
        filtered_temp_var_df = temp_var_df[(temp_var_df.FILE_NAME == req_file_name)]  # Filter data frame by check
        filtered_temp_var_df = filtered_temp_var_df.drop_duplicates(subset='VAR_NAME', keep='first')
        req_file.write('TEMPBEG\n')
        var_template = "{}   {};{}\n"

        for idx, var_info in filtered_temp_var_df.iterrows():
            var_name = var_info.loc['VAR_NAME']
            var_type = var_info.loc['TYPE']
            var_code = var_info.loc['CODE']
            req_file.write(var_template.format(var_name, var_type, var_code))

        req_file.write("TEMPEND\n\n")

        # Write CHECKS
        filtered_checks_df = checks_df[(checks_df.FILE_NAME == req_file_name)]  # Filter data frame by check
        check_line1_template = "{} {}{} {} {}{}{}\n"

        for idx, var_info in filtered_checks_df.iterrows():
            edit_check_def = var_info.loc['EDIT_CHECK_DEFINITION_LINE']
            edit_check_name = var_info.loc['CHECK_NAME']
            legal_illegal = var_info.loc['LEGAL_ILLEGAL']
            missing_records = var_info.loc['MISSING_RECORDS']
            error_code = var_info.loc['ERROR_CODE']
            check_description = var_info.loc['CHECK_DESCRIPTION']
            check_comment = var_info.loc['CHECK_COMMENT']
            padded_spaces_1 = " " * (15 - len(edit_check_name))
            padded_spaces_2 = " " * (50 - len(check_comment))
            padded_spaces_3 = " " * (6 - len(error_code))
            check_type = var_info.loc['TYPE']

            # write main check comment
            print(check_description)
            if type(check_description) != float:
                print(check_description, "in statment")
                check_description = tf.format_by_charater_length_with_keyword(max_length=50, text=check_description,
                                                                              keyword="COMMENT ")
                req_file.write(check_description)

                req_file.write("\n\n")

            # write line 1
            req_file.write(check_line1_template.format(edit_check_def, edit_check_name + padded_spaces_1,
                                                       legal_illegal, missing_records, error_code + padded_spaces_3,
                                                       check_comment + padded_spaces_2, check_type))

            # Write Code
            code_template = "{}\n\n"
            check_code = var_info.loc['CHECK_CODE']
            status = var_info.loc['STATUS']

            if status == "Ready to be generated using REQ file program":
                check_code = tf.format_by_charater_length_with_keyword(max_length=74, keyword="  ", text=check_code)
                req_file.write(code_template.format(check_code))
            else:
                print(status)
