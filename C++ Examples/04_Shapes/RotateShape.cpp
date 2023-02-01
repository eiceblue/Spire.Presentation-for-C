#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"RotateShape.pptx";
	std::wstring outputFile = OutputPath"RotateShape.pptx";

	//Load a PPT document
	Presentation* ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Get the shapes 
	IAutoShape* shape = dynamic_cast<IAutoShape*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(0));

	//Set the rotation
	shape->SetRotation(60);

	(dynamic_cast<IAutoShape*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(1)))->SetRotation(120);
	(dynamic_cast<IAutoShape*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(2)))->SetRotation(180);
	(dynamic_cast<IAutoShape*>(ppt->GetSlides()->GetItem(0)->GetShapes()->GetItem(3)))->SetRotation(240);

	//Save the document
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete ppt;
}
