def save_diff_to_docx(target_file_name, template_text, target_text, result_folder, fake_folder):
    doc = Document()
    doc.add_heading(f'Differences in {target_file_name}', level=1)

    diff_result = list(unified_diff(template_text.splitlines(), target_text.splitlines(), lineterm=''))
    
    # Calculate content matching percentage
    accuracy_content, _ = calculate_accuracy(template_text, target_text, True)

    # Add line-by-line differences to the document
    diff_paragraph = doc.add_paragraph()
    for line in diff_result:
        run = diff_paragraph.add_run(line)

        # Highlight added content in red
        if line.startswith('+'):
            font = run.font
            font.color.rgb = RGBColor(255, 0, 0)

    # Add table for differences
    if any(line.startswith('+') for line in diff_result):
        doc.add_heading('Differences in Table', level=2)
        table = doc.add_table(rows=1, cols=2)
        table.style = 'TableGrid'
        table.autofit = False
        table.columns[0].width = 200
        table.columns[1].width = 400

        for line in diff_result:
            if line.startswith('+'):
                cells = table.add_row().cells
                cells[0].text = 'Added'
                cells[1].text = line[1:]

                # Highlight added content based on matching percentage
                color = get_highlight_color(accuracy_content)
                cells[1].paragraphs[0].runs[0].font.color.rgb = color

    diff_docx_path = os.path.join(fake_folder, f'Differences_{target_file_name}.docx')
    doc.save(diff_docx_path)

def get_highlight_color(accuracy_content):
    if accuracy_content == 1.0:
        return RGBColor(0, 255, 0)  # Green color for 100% match
    elif accuracy_content >= 0.8:
        return RGBColor(255, 255, 0)  # Yellow color for 80% to 99% match
    else:
        return RGBColor(255, 0, 0)  # Red color for less than 80% match
