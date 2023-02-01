#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"BordersAndShading.pptx";
	std::wstring outputFile = OutputPath"BordersAndShading.pptx";

	//Load a PPT document
	Presentation* presentation = new Presentation();
	presentation->LoadFromFile(inputFile.c_str());

	IAutoShape* shape = dynamic_cast<IAutoShape*>(presentation->GetSlides()
		->GetItem(0)->GetShapes()->GetItem(0));

	//Set line color and width of the border
	shape->GetLine()->SetFillType(FillFormatType::Solid);
	shape->GetLine()->SetWidth(3);
	shape->GetLine()->GetSolidFillColor()->SetColor(Color::GetLightYellow());

	//Set the gradient fill color of shape

	shape->GetFill()->SetFillType(FillFormatType::Gradient);
	shape->GetFill()->GetGradient()->SetGradientShape(GradientShapeType::Linear);
	shape->GetFill()->GetGradient()->GetGradientStops()->Append(1.0f, KnownColors::LightBlue);
	shape->GetFill()->GetGradient()->GetGradientStops()->Append(0, KnownColors::LightSkyBlue);

	//Set the shadow for the shape
	//Spire::Presentation::Drawing::OuterShadowEffect* shadow = new Spire::Presentation::Drawing::OuterShadowEffect();
	OuterShadowEffect* shadow = new OuterShadowEffect();

	shadow->SetBlurRadius(20);
	shadow->SetDirection(30);
	shadow->SetDistance(8);
	shadow->GetColorFormat()->SetColor(Color::GetLightSeaGreen());
	shape->GetEffectDag()->SetOuterShadowEffect(shadow);

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2007);
	delete presentation;

}
