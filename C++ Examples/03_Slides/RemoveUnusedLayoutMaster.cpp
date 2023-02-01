#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"PPTSample_1.pptx";
	std::wstring outputFile = OutputPath"RemoveUnusedLayoutMaster.pptx";

	Presentation* ppt = new Presentation();
	ppt->LoadFromFile(inputFile.c_str());

	//Create an array list
	std::vector<IActiveSlide*> list;
	for (int i = 0; i < ppt->GetSlides()->GetCount(); i++)
	{
		//Get the layout used by slide
		IActiveSlide* layout = dynamic_cast<IActiveSlide*>(ppt->GetSlides()->GetItem(i)->GetLayout());
		list.push_back(layout);
	}

	//Loop through masters and layouts
	for (int i = 0; i < ppt->GetMasters()->GetCount(); i++)
	{
		IMasterLayouts* masterlayouts = ppt->GetMasters()->GetItem(i)->GetLayouts();
		for (int j = masterlayouts->GetCount() - 1; j >= 0; j--)
		{
			if (!(std::find(list.begin(), list.end(), dynamic_cast<IActiveSlide*>(masterlayouts->GetItem(j))) != list.end()))
			{
				//Remove unused layout
				masterlayouts->RemoveMasterLayout(j);
			}
		}

	}
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);
	delete ppt;
}
