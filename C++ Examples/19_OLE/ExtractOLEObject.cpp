#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ExtractOLEObject.pptx";
	std::wstring outputFile_px = OutputPath"ExtractOLEObject.pptx";
	std::wstring outputFile_p = OutputPath"ExtractOLEObject.ppt";
	std::wstring outputFile_xls = OutputPath"tractOLEObject.xls";
	std::wstring outputFile_xlsx = OutputPath"ExtractOLEObject.xlsx";
	std::wstring outputFile_doc = OutputPath"ExtractOLEObject.doc";
	std::wstring outputFile_docx = OutputPath"ExtractOLEObject.docx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Loop through the slides and shapes
	for (int i = 0; i < presentation->GetSlides()->GetCount(); i++)
	{
		ISlide* slide = presentation->GetSlides()->GetItem(i);
		for (int k = 0; k < slide->GetShapes()->GetCount(); k++)
		{
			IShape* shape = slide->GetShapes()->GetItem(k);
			if (dynamic_cast<Spire::Presentation::IOleObject*>(shape) != nullptr)
			{
				//Find OLE object
				Spire::Presentation::IOleObject* oleObject = dynamic_cast<Spire::Presentation::IOleObject*>(shape);

				//Get its data and write to file
				Stream* stream = oleObject->GetDataStream();

				if (wcscmp(oleObject->GetProgId(), L"Excel.Sheet.8") == 0)
				{
					stream->Save(outputFile_xls.c_str());
				}
				//ORIGINAL LINE: case "Excel.Sheet.12":
				else if (wcscmp(oleObject->GetProgId(), L"Excel.Sheet.12") == 0)
				{
					stream->Save(outputFile_xlsx.c_str());
				}
				//ORIGINAL LINE: case "Word.Document.8":
				else if (wcscmp(oleObject->GetProgId(), L"Word.Document.8") == 0)
				{
					stream->Save(outputFile_doc.c_str());
				}
				//ORIGINAL LINE: case "Word.Document.12":
				else if (wcscmp(oleObject->GetProgId(), L"Word.Document.12") == 0)
				{
					stream->Save(outputFile_docx.c_str());
				}
				//ORIGINAL LINE: case "PowerPoint.Show.8":
				else if (wcscmp(oleObject->GetProgId(), L"PowerPoint.Show.8") == 0)
				{
					stream->Save(outputFile_p.c_str());
				}
				//ORIGINAL LINE: case "PowerPoint.Show.12":
				else if (wcscmp(oleObject->GetProgId(), L"PowerPoint.Show.12") == 0)
				{
					stream->Save(outputFile_px.c_str());
				}
				stream->Close();
			}
		}
	}
	delete presentation;

}
