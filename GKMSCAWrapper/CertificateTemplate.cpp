#include "stdafx.h"
#include "CertificateTemplate.h"

#import "../dependencies/scrdenrl.dll" no_namespace

#define SCARD_ENROLL_CA_MACHINE_NAME 1

#define SCARD_ENROLL_USER_CERT_TEMPLATE 1
#define SCARD_ENROLL_MACHINE_CERT_TEMPLATE 2

#define SCARD_ENROLL_TEMPLATE_INFO_OID 3

namespace GK {
	namespace MSCAWrapper {
		CertificateTemplate::CertificateTemplate(System::String ^templateName, System::String ^templateOID, TemplateCategory templateCategory)
			: TemplateType(templateCategory)
		{
			this->TemplateName = templateName;
			this->TemplateOID = templateOID;
		}

		array<CertificateTemplate^>^ GK::MSCAWrapper::CertificateTemplate::RetrieveCertificateTemplates(IWebEnrlServer *pWebEnrlServerEnroll, int typeflag4Enumeration, TemplateCategory templateCategory)
		{
			long lngTemplateCount = pWebEnrlServerEnroll->getCertTemplateCount(typeflag4Enumeration);

			array<CertificateTemplate^> ^returnedTemplates = gcnew array<CertificateTemplate ^>(lngTemplateCount);

			for (int i = 0; i < lngTemplateCount; ++i)
			{
				System::IntPtr pBstrTemplateName = System::IntPtr::Zero;
				VARIANT varOID;

				try
				{
					BSTR bsCurrentTemplateName = pWebEnrlServerEnroll->enumCertTemplateName(i, typeflag4Enumeration);
					System::IntPtr pBstrTemplateName = System::IntPtr(bsCurrentTemplateName);
					System::String ^sCurrentTemplateName = System::Runtime::InteropServices::Marshal::PtrToStringBSTR(pBstrTemplateName);

					varOID = pWebEnrlServerEnroll->getCertTemplateInfo(bsCurrentTemplateName, SCARD_ENROLL_TEMPLATE_INFO_OID);
					System::String ^sCurrentTemplateOID = (System::String ^)System::Runtime::InteropServices::Marshal::GetObjectForNativeVariant(System::IntPtr(&varOID));

					returnedTemplates[i] = gcnew CertificateTemplate(sCurrentTemplateName, sCurrentTemplateOID, templateCategory);
				}
				finally
				{
					if (System::IntPtr::Zero != pBstrTemplateName)
						System::Runtime::InteropServices::Marshal::FreeBSTR(pBstrTemplateName);
					VariantClear(&varOID);
				}
			}

			return returnedTemplates;
		}

		array<CertificateTemplate^>^ GK::MSCAWrapper::CertificateTemplate::RetrieveAllUserCertificateTemplates()
		{
			HRESULT hr;
			IWebEnrlServer *pWebEnrlServerEnroll = NULL;

			try
			{
				hr = CoCreateInstance(
					__uuidof(WebEnrlServer),
					NULL,
					CLSCTX_INPROC_SERVER,
					__uuidof(IWebEnrlServer),
					(void **)&pWebEnrlServerEnroll);
				ASSERT_COM_SUCCESS(hr, "Could not create WebEnrlServer object");

				return RetrieveCertificateTemplates(pWebEnrlServerEnroll, SCARD_ENROLL_USER_CERT_TEMPLATE, TemplateCategory::USER);
			}
			finally
			{
				CHECK_COM_NULL_AND_RELEASE(pWebEnrlServerEnroll);
			}
		}

		array<CertificateTemplate^>^ GK::MSCAWrapper::CertificateTemplate::RetrieveAllMachineCertificateTemplates()
		{
			HRESULT hr;
			IWebEnrlServer *pWebEnrlServerEnroll = NULL;

			try
			{
				hr = CoCreateInstance(
					__uuidof(WebEnrlServer),
					NULL,
					CLSCTX_INPROC_SERVER,
					__uuidof(IWebEnrlServer),
					(void **)&pWebEnrlServerEnroll);
				ASSERT_COM_SUCCESS(hr, "Could not create WebEnrlServer object");

				return RetrieveCertificateTemplates(pWebEnrlServerEnroll, SCARD_ENROLL_MACHINE_CERT_TEMPLATE, TemplateCategory::MACHINE);
			}
			finally
			{
				CHECK_COM_NULL_AND_RELEASE(pWebEnrlServerEnroll);
			}
		}
	}
}