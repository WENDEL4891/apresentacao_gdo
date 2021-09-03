from pptx import Presentation


def replace_text(path_name_source_file, path_name_result_file, replace_dict):
    prs = Presentation(path_name_source_file)
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for match, replacement in replace_dict.items():                
                if match in shape.text:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            p = paragraph._p  # the lxml element containing the `<a:p>` paragraph element
        #                            remove all but the first run
                            for idx, run in enumerate(paragraph.runs):
                                if idx == 0:
                                    continue
                                p.remove(run._r)
                            cur_text = run.text
                            new_text = cur_text.replace(str(match), str(replacement))
                            run.text = new_text
    prs.save(path_name_result_file)


def replace_text_table_header(path_name_source_file, path_name_result_file, replace_dict):
    prs = Presentation(path_name_source_file)
    for slide in prs.slides:
        for shape in slide.shapes:
                for match, replacement in replace_dict.items():
                    if shape.has_table:
                        print(dir(shape.table.rows))
#                         for cell in shape.table.iter_cells():
#                         for cell in shape.table.rows:
#                             print(cell.text)
#                             if match == cell.text:
#                                 for paragraph in cell.text_frame.paragraphs:
#                                    for run in paragraph.runs:
#                                        p = paragraph._p  # the lxml element containing the `<a:p>` paragraph element
#         #                            remove all but the first run
#                                        for idx, run in enumerate(paragraph.runs):
#                                            if idx == 0:
#                                                continue
#                                            p.remove(run._r)
#                                        cur_text = run.text
#                                        new_text = cur_text.replace(str(match), str(replacement))
#                                        run.text = new_text
#     prs.save(path_name_result_file)