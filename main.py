import win32com.client
import os

ea_repo = win32com.client.Dispatch("EA.Repository")
print(ea_repo)
ea_repo.OpenFile("D:\\Code\\EA\\Test.qea")
for i in range(ea_repo.Models.Count):
    package = ea_repo.Models.GetAt(i)
    print(f"Package Name: {package.Name}, Package ID: {package.PackageID}, GUID: {package.PackageGUID}")

package_id = 1  # Replace with the actual package ID
package = ea_repo.GetPackageByID(package_id)
output_path = "D:\\Code\\EA\\test"
ea_repo.GetProjectInterface().RunHTMLReport(package.PackageGUID, output_path, "GIF", "<default>", ".html")

ea_repo.CloseFile()
ea_repo.Exit()