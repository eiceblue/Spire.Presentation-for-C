#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"InputTemplate.pptx";
	std::wstring outputFile = OutputPath"LoadFromStream.pptx";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	Stream* stream = new Stream(inputFile.c_str());
	ppt->LoadFromStream(stream, FileFormat::Pptx2013);

	//Save to file.
	ppt->SaveToFile(outputFile.c_str(), FileFormat::Pptx2013);

	stream->Dispose();
	delete ppt;

}
