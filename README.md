# Spire.Presentation for C++ - A C++ Library for Processing PowerPoint Documents

[![Foo](https://i.imgur.com/m1lcAIy.png)](https://www.e-iceblue.com/Introduce/presentation-for-CPP.html)

[Product Page](https://www.e-iceblue.com/Introduce/presentation-for-CPP.html) | [Forum](https://www.e-iceblue.com/forum/spire-presentation-f14.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html) 

[Spire.Presentation for C++](https://www.e-iceblue.com/Introduce/presentation-for-CPP.html) is a professional **PowerPoint C++ API** that enables developers to create, read, write, modify, and convert PowerPoint documents on any C++ platforms without installing Microsoft PowerPoint.

This API supports PPT, PPS, PPTX and PPSX presentation formats. It provides functions such as managing text, image, shapes, tables, animations, audio, and video on slides and supports exporting presentation slides to JPG, PDF, XPS, SVG, HTML and other formats.

### 100% Standalone C++ API

Spire.Presentation for C++ is a totally independent C++ PowerPoint API which doesn't require Microsoft PowerPoint to be installed on system.

### Freely Operate PowerPoint Files

- Create/Save/Merge/Split/Print PowerPoint Document.
- Protect/Unprotect PowerPoint Document.
- Create/Add/Delete/Hide/Show/Move slides.
- Add/Remove/Extract/Replace comments and notes in PowerPoint.
- Add/Remove/Revise/Extract/Replace texts and images in PowerPoint.
- Work with charts, tables and SmartArt in PowerPoint.
- Insert/Modify/Remove hyperlinks.
- Add/Remove text and image watermark.
- Insert/Replace/Extract Audio and Video.

### Powerful & High Quality PowerPoint File Conversion

- Convert PowerPoint to HTML
- Convert PowerPoint to XPS
- Convert PowerPoint to SVG
- Convert PowerPoint to PDF
- Convert PowerPoint to PPTX
- Convert PowerPoint to Image/Image to PowerPoint 

### Examples

### Convert PowerPoint to PDF in C++

```c++
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ToPDF.pptx";
	std::wstring outputFile = OutputPath"ToPDF.pdf";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Save the document to PDF format
	ppt->SaveToFile(outputFile.c_str(), FileFormat::PDF);

	delete ppt;
}
```

### Convert PowerPoint to Images in C++

```c++
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"ToImage.pptx";
	std::wstring outputFile = OutputPath"Image/ToImage/";

	//Create a PPT document
	Presentation* ppt = new Presentation();
	//Load the file from disk.
	ppt->LoadFromFile(inputFile.c_str());

	//Save PPT document to images
	SlideCollection* slides = ppt->GetSlides();
	for (int i = 0; i < slides->GetCount(); i++)
	{
		ISlide* slide = slides->GetItem(i);
		Stream* image = slide->SaveAsImage();
		image->Save((outputFile + L"ToImage_img_" + to_wstring(i) + L".png").c_str());
	}

	delete ppt;
}
```

### Encrypt PowerPoint in C++

```c++
#include "pch.h"

using namespace std;
using namespace Spire::Presentation;

int main()
{
	std::wstring inputFile = DataPath"Encrypt.pptx";
	std::wstring outputFile = OutputPath"Encrypt.pptx";

	//Create a PPT document
	Presentation* presentation = new Presentation();

	//Load the document from disk
	presentation->LoadFromFile(inputFile.c_str());

	//Get the password that the user entered
	std::wstring password = L"e-iceblue";

	//Encrypy the document with the password
	presentation->Encrypt(password.c_str());

	//Save the document
	presentation->SaveToFile(outputFile.c_str(), FileFormat::Pptx2010);
	delete presentation;

}
```

[Product Page](https://www.e-iceblue.com/Introduce/presentation-for-CPP.html) | [Forum](https://www.e-iceblue.com/forum/spire-presentation-f14.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html) 

