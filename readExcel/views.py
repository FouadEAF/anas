import os
import json
import win32com.client
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
import pythoncom


@csrf_exempt
def run_excel(request):
    if request.method == "POST":
        try:
            # Parse JSON data from the request body
            data = json.loads(request.body)
            # Ensure "ranges" is available
            ranges = data.get("ranges", {})
            print("ranges from frontend:", ranges)

            # Construct the absolute path to the Excel file
            base_dir = os.path.dirname(os.path.abspath(__file__))
            file_path = os.path.join(base_dir, "assets", "Training.xlsx")
            solver_addin_path = os.path.join(base_dir, "assets", "SOLVER.XLAM")
            print("Excel file path:", file_path)

            # Check if file exists
            if not os.path.exists(file_path):
                return JsonResponse({"success": False, "message": "Excel file not found"}, status=404)

            # Check if the Solver add-in exists
            if not os.path.exists(solver_addin_path):
                return JsonResponse({"success": False, "message": "Solver add-in not found"}, status=404)

            # Initialize COM
            pythoncom.CoInitialize()

            # Initialize Excel Application (requires Windows with Excel installed)
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False  # Keep Excel hidden during the operation

            # Open the Excel file
            workbook = excel.Workbooks.Open(file_path)
            sheet = workbook.Sheets(1)  # Access the first sheet

            # Update the cells as per the frontend data
            print("Update cells based on the ranges provided")
            for cell, value in ranges.items():
                print(f"Setting {cell} to {value}")
                sheet.Range(cell).Value = value  # Update the cell value

            # Force Excel to calculate the formulas
            print("Force Excel to calculate formulas")
            excel.Calculate()

            # Get the calculated result from cell D40
            result_d40 = sheet.Range("D40").Value
            print("Result from D40:", result_d40)

            # Save the workbook after making the changes
            workbook.Save()

            # Close Excel application
            workbook.Close()
            excel.Quit()

        except json.JSONDecodeError:
            return JsonResponse({"success": False, "message": "Invalid JSON data"}, status=400)
        except Exception as e:
            return JsonResponse({"success": False, "error": str(e)}, status=400)

        return JsonResponse({"success": True, "data": result_d40}, status=200)

    return JsonResponse({"success": False, "message": "Invalid request method"}, status=405)

# import os
# import json
# import pythoncom
# import win32com.client
# from django.http import JsonResponse
# from django.views.decorators.csrf import csrf_exempt


# @csrf_exempt
# def run_excel(request):
#     if request.method == "POST":
#         try:
#             # Initialize COM
#             pythoncom.CoInitialize()  # Initialize the COM library

#             # Parse the incoming data
#             data = json.loads(request.body)
#             ranges = data.get("ranges", {})  # Ensure "ranges" is available

#             # Construct the path to the Excel file and Solver add-in
#             base_dir = os.path.dirname(os.path.abspath(
#                 __file__))  # Project root directory
#             excel_file_path = os.path.join(
#                 base_dir, "assets", "Training.xlsx")  # Excel file with VBA code
#             solver_addin_path = os.path.join(
#                 base_dir, "assets", "SOLVER.XLAM")  # Solver add-in

#             # Check if the Excel file exists
#             if not os.path.exists(excel_file_path):
#                 return JsonResponse({"success": False, "message": "Excel file not found"}, status=404)

#             # Check if the Solver add-in exists
#             if not os.path.exists(solver_addin_path):
#                 return JsonResponse({"success": False, "message": "Solver add-in not found"}, status=404)

#             # Initialize Excel Application (requires Windows with Excel installed)
#             excel = win32com.client.Dispatch("Excel.Application")
#             excel.Visible = True  # Set to True if you want Excel to be visible during execution

#             # Open the Excel file
#             print("Opening Excel workbook...")
#             workbook = excel.Workbooks.Open(excel_file_path)
#             print("Workbook opened successfully")

#             # Load the Solver add-in
#             # excel.Application.AddIns.Add(solver_addin_path)
#             print("Loading Solver add-in...")
#             # Ensure Solver is installed
#             # excel.Application.AddIns("Solver").Installed = True
#             print("Solver add-in loaded successfully")

#             # Update the cell values as provided in the frontend data
#             sheet = workbook.Sheets(1)  # Access the first sheet
#             for cell, value in ranges.items():
#                 # Set value for each specified cell
#                 print(f"Setting cell {cell} to {value}")
#                 sheet.Range(cell).Value = value

#             # Save the workbook after updating values
#             workbook.Save()

#             # Try to run the macro and catch any errors
#             # try:
#             #     print("Running macro 'RunSolver'...")
#             #     excel.Application.Run("RunSolver")
#             #     print("Macro 'RunSolver' executed successfully")
#             # except Exception as e:
#             #     return JsonResponse({"success": False, "error": f"Macro failed: {str(e)}"}, status=400)

#             # Get the result from cell D40 (assuming it is calculated by the Solver macro)
#             result_d40 = sheet.Range("D40").Value
#             print("Result from D40:", result_d40)

#             # Save and close the workbook
#             workbook.Close()

#             # Close Excel application
#             excel.Quit()

#         except Exception as e:
#             return JsonResponse({"success": False, "error": f"Error occurred: {str(e)}"}, status=400)

#         # Return the value of D40
#         return JsonResponse({"success": True, "result": result_d40}, status=200)

#     return JsonResponse({"success": False, "message": "Invalid request method"}, status=405)


# import openpyxl
# from django.http import JsonResponse
# from django.views.decorators.csrf import csrf_exempt
# import json
# import os


# @csrf_exempt
# def run_excel(request):
#     if request.method == "POST":
#         try:
#             # Parse JSON data from the request body
#             data = json.loads(request.body)
#             # Ensure "ranges" is available
#             ranges = data.get("ranges", {})
#             print("ranges from frontend:", ranges)

#             # Construct the absolute path to the Excel file
#             # Get current file's directory  # Path to BLEVE.xlsx
#             base_dir = os.path.dirname(os.path.abspath(__file__))
#             file_path = os.path.join(base_dir, "assets", "Training.xlsx")
#             solver_addin_path = os.path.join(
#                 base_dir, "assets", "SOLVER.XLAM")  # Solver add-in
#             print("Excel file path:", file_path)

#             print("Check if file exists")
#             # Check if file exists
#             if not os.path.exists(file_path):
#                 return JsonResponse({"success": False, "message": "Excel file not found"}, status=404)

#              # Check if the Solver add-in exists
#             if not os.path.exists(solver_addin_path):
#                 return JsonResponse({"success": False, "message": "Solver add-in not found"}, status=404)

#             # Load the Excel file
#             print("Load the Excel file")
#             wb = openpyxl.load_workbook(file_path)
#             sheet = wb.active

#             print("Update cells based on the ranges provided")
#             # Update cells based on the ranges provided
#             for cell, value in ranges.items():
#                 sheet[cell] = value  # Update each specified cell

#             print("Save changes back to the file")
#             # Save changes back to the file
#             wb.save(file_path)

#             print("Get the result from cell D40")
#             # Get the result from cell D40
#             result_d40 = sheet["D40"].value
#             print("result", result_d40)

#         except json.JSONDecodeError:
#             return JsonResponse({"success": False, "message": "Invalid JSON data"}, status=400)
#         except Exception as e:
#             return JsonResponse({"success": False, "error": str(e)}, status=400)
#         # Return the value of D40
#         return JsonResponse({"success": True, "data": result_d40}, status=200)
#     return JsonResponse({"success": False, "message": "Invalid request method"}, status=405)
