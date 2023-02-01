#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Bullets.pptx";
	std::wstring outputFile = OutputPath"Bullets.pptx";

	//Load a PPT document
	Presentation* presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	IAutoShape* shape = dynamic_cast<IAutoShape*>(presentation->GetSlides()->GetItem(0)->GetShapes()->GetItem(1));

	for (int t = 0; t < shape->GetTextFrame()->GetParagraphs()->GetCount(); t++)
	{
		TextParagraph* para = shape->GetTextFrame()->GetParagraphs()->GetItem(t);
		//Add the bullets
		para->SetBulletType(TextBulletType::Numbered);
		para->SetBulletStyle(NumberedBulletStyle::BulletRomanLCPeriod);
	}

	//Save the document and launch
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;

}
