import os
import pythoncom
import win32com.client


def ungroup_shapes_in_ppt(input_path, output_path):
    """
    Ungroups all grouped shapes (including nested) in a PowerPoint presentation.
    Saves the modified file to the specified output path.
    """
    pythoncom.CoInitialize()
    ppt = win32com.client.Dispatch("PowerPoint.Application")

    try:
        # Try setting PowerPoint visibility to False (silent mode)
        try:
            ppt.Visible = False
        except Exception:
            pass  # Some systems may not allow hiding PowerPoint

        input_path = os.path.abspath(input_path)
        output_path = os.path.abspath(output_path)

        presentation = ppt.Presentations.Open(
            input_path, ReadOnly=0, Untitled=0, WithWindow=0
        )

        for slide in presentation.Slides:
            ungrouped = True
            # Keep ungrouping until no more groups are found (recursive)
            while ungrouped:
                ungrouped = False
                for i in range(slide.Shapes.Count, 0, -1):
                    shape = slide.Shapes(i)
                    if shape.Type == 6:  # msoGroup = 6
                        try:
                            shape.Ungroup()
                            ungrouped = True
                        except Exception:
                            continue  # Safe skip for ungroupable objects

        presentation.SaveAs(output_path)
        presentation.Close()

    except Exception as e:
        print(f"[Ungroup Error] Failed: {str(e)}")

    finally:
        ppt.Quit()
        pythoncom.CoUninitialize()
