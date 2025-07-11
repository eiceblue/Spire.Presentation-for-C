
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

string  wstring2string(const std::wstring& wstr)
{
	std::string result;
	result.reserve(wstr.size());
	for (size_t i = 0; i < wstr.size(); ++i)
	{
		result += static_cast<char>(wstr[i] & 0xFF);
	}
	return result;
}

int main()
{
	wstring inputFile = DATAPATH"SaveChartAsImage.pptx";
	wstring outputFile = OUTPUTPATH"Image/SaveChartAsImage.png";

	//Create a PPT document
	intrusive_ptr<Presentation> ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Save chart as image in .png format
	intrusive_ptr<Stream> image = ppt->GetSlides()->GetItem(0)->GetShapes()->SaveAsImage(0);
	std::ofstream output(wstring2string(outputFile), std::ios::binary);
	image->Save(output);
	output.flush();
	output.close();
	image->Close();
}
