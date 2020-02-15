﻿// EditorFactory.h

#pragma once

#include "EditorDocument.h"

/***************************************************************************
VSPackage1.pkgdef contains:


[$RootKey$\Editors\{b71aef19-43b7-4539-ba86-6f4823b1d34c}]
@="VSPackage1"
"Package"="{31ed66c3-253a-4fb7-9788-acff1bbf27b4}"
"DisplayName"="VSPackage1"
"ExcludeDefTextEditor"=dword:00000001
"AcceptBinaryFiles"=dword:00000000

[$RootKey$\Editors\{b71aef19-43b7-4539-ba86-6f4823b1d34c}\LogicalViews]
"{7651A703-06E5-11D1-8EBD-00A0C90F26EA}"=""

[$RootKey$\Editors\{b71aef19-43b7-4539-ba86-6f4823b1d34c}\Extensions]
"vspackage1"=dword:00000032

which informs the shell that VSPackage1Package is the package to create to have this editor 
registered with Visual Studio.  When VSPackage1Package is sited by Visual Studio, 
VSPackage1Package::PostSited creates an instance of EditorFactory and then queries
for SID_SVsRegisterEditors and then calls IVsRegisterEditors->RegisterEditor to register
the new instance of EditorFactory with Visual Studio.

Visual Studio then uses the EditorFactory instance to create EditorDocument
instances, which is the actual editor instance.
***************************************************************************/

class EditorFactory :
	// Use ATL to take care of common COM infrastructure
	public CComObjectRootEx<CComSingleThreadModel>,
	// IVsEditorFactory is required to be an editor factory
	public IVsEditorFactoryImpl<EditorFactory>
{

// Provides a portion of the implementation of IUnknown, in particular the list of interfaces
// the EditorFactory object will support via QueryInterface
BEGIN_COM_MAP(EditorFactory)
    COM_INTERFACE_ENTRY(IVsEditorFactory)
END_COM_MAP()

// COM objects typically should not be cloned, and this prevents cloning by declaring the 
// copy constructor and assignment operator private (NOTE:  this macro includes the declaration of
// a private section, so everything following this macro and preceding a public or protected 
// section will be private).
VSL_DECLARE_NOT_COPYABLE(EditorFactory)

protected:
	EditorFactory()
	{
	}

	virtual ~EditorFactory()
	{
	}

public:

#pragma warning(push)
#pragma warning(disable : 4480) // // warning C4480: nonstandard extension used: specifying underlying type for enum
	enum PhysicalViewId : unsigned int
	{
		Unsupported,
		Primary
	};
#pragma warning(pop)

	PhysicalViewId GetPhysicalViewId(REFGUID rguidLogicalView)
	{
		if(LOGVIEWID_Primary == rguidLogicalView || LOGVIEWID_TextView == rguidLogicalView)
		{
			return Primary;
		}

		return Unsupported;
	}

	BSTR GetPhysicalViewBSTR(PhysicalViewId viewId)
	{
		if(Primary == viewId)
		{
			// Note that the value of LOGVIEWID_TextView in the registry is empty string.
			// The values under LogicalViews for an editor should match up to the values 
			// returned by this method, so that GetPhysicalViewId below can properly
			// map the strings back to a PhysicalViewId.
			return ::SysAllocString(L"");
		}

		return NULL;
	}

	PhysicalViewId GetPhysicalViewId(LPCOLESTR szPhysicalView)
	{
		if(szPhysicalView != NULL && szPhysicalView[0] == L'\0')
		{
			return Primary;
		}

		return Unsupported;
	}

	bool CanShareBuffer(PhysicalViewId /*physicalViewId*/)
	{
		return false;
	}

	void CreateSingleViewObject(
		PhysicalViewId physicalViewId, 
		CComPtr<IUnknown>& rspViewObject, 
		CComBSTR& rbstrEditorCaption, 
		const GUID*& rpguidCommandUI, 
		VSEDITORCREATEDOCWIN& /*rCreateDocumentWindowUI*/)
	{
		if(physicalViewId == Primary)
		{
			CComObject<EditorDocument<> > *pDocument;
			VSL_CHECKHRESULT(CComObject<EditorDocument<> >::CreateInstance(&pDocument));
			HRESULT hr = pDocument->QueryInterface(&rspViewObject);
			if(FAILED(hr))
			{
				// If QueryInterface failed, then there is something wrong with the object.
				// Delete it and throw an exception for the error.
				delete pDocument;
				VSL_CREATE_ERROR_HRESULT(hr);
			}

			// NOTE - if the file is read only [Read Only] will be appended to the caption
			rbstrEditorCaption.LoadStringW(IDS_EDITORCAPTION);

			rpguidCommandUI = &CLSID_VSPackage1EditorDocument;

			// pCreateDocumentWindowUI will be initialized to 0, and see we provide
			// no GUI for the user to cancel from so leave it 0

			return;
		}
		ERRHR(E_FAIL); // This should never happen
	}

private:

};
