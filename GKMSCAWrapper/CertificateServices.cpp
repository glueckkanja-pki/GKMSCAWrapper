#include "StdAfx.h"
#include "CertificateServices.h"
#include <Certsrv.h>
#include "SingleValueRestriction.h"
#include "NoRestriction.h"

#pragma comment(lib, "ole32")		// COM
#pragma comment(lib, "oleaut32")	// VARIANTs
#pragma comment(lib, "Certidl")		// Certificate Services COM classes
#pragma comment(lib, "Certadm")		// ditto

namespace GK {
namespace MSCAWrapper {

	CertificateServices::CertificateServices(System::String ^caConfiguration)
		: MSCAObject(caConfiguration)
	{
	}

	void CertificateServices::revokeCertificate(System::String ^strSerial, RevocationReason reason, System::DateTime date)
	{
		HRESULT hr;
		ICertAdmin2 *pCertAdmin = NULL;

		try
		{
			hr = CoCreateInstance(CLSID_CCertAdmin,
			  NULL,
			  CLSCTX_INPROC_SERVER,
			  IID_ICertAdmin2,
			  (void **)&pCertAdmin);
			ASSERT_COM_SUCCESS(hr, "Could not create CCertAdmin object");
			
			BSTR bsCaConfiguration = NULL;
			System::IntPtr pBstrCAConfig = System::Runtime::InteropServices::Marshal::StringToBSTR(caConfiguration);
			bsCaConfiguration = reinterpret_cast<BSTR>(pBstrCAConfig.ToPointer());

			BSTR bsSerial = NULL;
			System::IntPtr pBstrSerial = System::Runtime::InteropServices::Marshal::StringToBSTR(strSerial);
			bsSerial = reinterpret_cast<BSTR>(pBstrSerial.ToPointer());

			hr = pCertAdmin->RevokeCertificate(bsCaConfiguration, bsSerial, System::Convert::ToInt32(reason), date.ToOADate());

			System::Runtime::InteropServices::Marshal::FreeBSTR(pBstrCAConfig);
			bsCaConfiguration = NULL;

			System::Runtime::InteropServices::Marshal::FreeBSTR(pBstrSerial);
			bsSerial = NULL;

			ASSERT_COM_SUCCESS(hr, "Could not revoke a certificate");
		}
		finally
		{
			CHECK_COM_NULL_AND_RELEASE(pCertAdmin);
		}
	}

	void CertificateServices::deleteCertificateRow(int requestID)
	{
		HRESULT hr;
		ICertAdmin2 *pCertAdmin = NULL;

		try
		{
			hr = CoCreateInstance(CLSID_CCertAdmin,
			  NULL,
			  CLSCTX_INPROC_SERVER,
			  IID_ICertAdmin2,
			  (void **)&pCertAdmin);
			ASSERT_COM_SUCCESS(hr, "Could not create CCertAdmin object");
			
			BSTR bsCaConfiguration = NULL;
			System::IntPtr pBstrCAConfig = System::Runtime::InteropServices::Marshal::StringToBSTR(caConfiguration);
			bsCaConfiguration = reinterpret_cast<BSTR>(pBstrCAConfig.ToPointer());

			LONG numDeleted;
			hr = pCertAdmin->DeleteRow(bsCaConfiguration, 0, 0, CVRC_TABLE_REQCERT, requestID, &numDeleted);

			System::Runtime::InteropServices::Marshal::FreeBSTR(pBstrCAConfig);
			bsCaConfiguration = NULL;

			ASSERT_COM_SUCCESS(hr, "Could not delete a certificate database row");

			if (numDeleted != 1)
				throw gcnew System::ArgumentException("Instead of exactly 1 row, " + numDeleted.ToString() + " rows were deleted", "requestID");
		}
		finally
		{
			CHECK_COM_NULL_AND_RELEASE(pCertAdmin);
		}
	}

	int CertificateServices::importCertificate(System::String ^sCertificate)
	{
		HRESULT hr;
		ICertAdmin2 *pCertAdmin = NULL;

		try
		{
			hr = CoCreateInstance(CLSID_CCertAdmin,
			  NULL,
			  CLSCTX_INPROC_SERVER,
			  IID_ICertAdmin2,
			  (void **)&pCertAdmin);
			ASSERT_COM_SUCCESS(hr, "Could not create CCertAdmin object");
			
			BSTR bsCaConfiguration = NULL;
			System::IntPtr pBstrCAConfig = System::Runtime::InteropServices::Marshal::StringToBSTR(caConfiguration);
			bsCaConfiguration = reinterpret_cast<BSTR>(pBstrCAConfig.ToPointer());

			BSTR bsCertificate = NULL;
			System::IntPtr pBstrCertificate = System::Runtime::InteropServices::Marshal::StringToBSTR(sCertificate);
			bsCertificate = reinterpret_cast<BSTR>(pBstrCertificate.ToPointer());
			
			LONG numRequest;

			hr = pCertAdmin->ImportCertificate(bsCaConfiguration, bsCertificate, CR_IN_BASE64HEADER, &numRequest);

			System::Runtime::InteropServices::Marshal::FreeBSTR(pBstrCAConfig);
			bsCaConfiguration = NULL;

			System::Runtime::InteropServices::Marshal::FreeBSTR(pBstrCertificate);
			bsCertificate = NULL;

			ASSERT_COM_SUCCESS(hr, "Could not import certificate");

			return numRequest;
		}
		finally
		{
			CHECK_COM_NULL_AND_RELEASE(pCertAdmin);
		}
	}

	int CertificateServices::importCertificate(array<System::Byte> ^baCertificate)
	{
		HRESULT hr;
		ICertAdmin2 *pCertAdmin = NULL;

		try
		{
			hr = CoCreateInstance(CLSID_CCertAdmin,
			  NULL,
			  CLSCTX_INPROC_SERVER,
			  IID_ICertAdmin2,
			  (void **)&pCertAdmin);
			ASSERT_COM_SUCCESS(hr, "Could not create CCertAdmin object");
			
			BSTR bsCaConfiguration = NULL;
			System::IntPtr pBstrCAConfig = System::Runtime::InteropServices::Marshal::StringToBSTR(caConfiguration);
			bsCaConfiguration = reinterpret_cast<BSTR>(pBstrCAConfig.ToPointer());

			char *binCertificate;
			NETBYTEARRAY_2_CHARSTRING(baCertificate, binCertificate);
			
			LONG numRequest;

			hr = pCertAdmin->ImportCertificate(bsCaConfiguration, reinterpret_cast<BSTR>(binCertificate), CR_IN_BASE64HEADER, &numRequest);

			System::Runtime::InteropServices::Marshal::FreeBSTR(pBstrCAConfig);
			bsCaConfiguration = NULL;

			delete[] binCertificate;

			ASSERT_COM_SUCCESS(hr, "Could not import certificate");

			return numRequest;
		}
		finally
		{
			CHECK_COM_NULL_AND_RELEASE(pCertAdmin);
		}
	}


	System::Collections::Generic::IList<CertColumn ^> ^CertificateServices::columns::get()
	{
		if (nullptr != _columns)
			return _columns;

		ICertView *pCertView = NULL;
		IEnumCERTVIEWCOLUMN *pEnumColumn = NULL;
		HRESULT hr;

		System::Collections::Generic::List<CertColumn ^> ^retArray = gcnew System::Collections::Generic::List<CertColumn ^>();

		try
		{
			pCertView = openCAConnection();

			hr = pCertView->EnumCertViewColumn(CVRC_COLUMN_SCHEMA, &pEnumColumn);
			ASSERT_COM_SUCCESS(hr, "Could not enumerate columns in CA database");

			hr = pEnumColumn->Reset();
			ASSERT_COM_SUCCESS(hr, "Could not reset column enumeration");

			LONG lngCurrentIndex = 0;
			while (S_OK == (hr = pEnumColumn->Next(&lngCurrentIndex)))
			{
				BSTR bstrName = NULL;
				hr = pEnumColumn->GetName(&bstrName);
				System::String ^msgError = System::String::Concat("Could not get the name of column #", lngCurrentIndex.ToString());
				ASSERT_COM_SUCCESS(hr, msgError);

				retArray->Add
				(
					gcnew CertColumn(
						caConfiguration, 
						System::Runtime::InteropServices::Marshal::PtrToStringBSTR(System::IntPtr(bstrName))
					)
				);
				System::Runtime::InteropServices::Marshal::FreeBSTR(System::IntPtr(bstrName));
			}

			if (S_FALSE != hr)		// S_FALSE = enumeration ends without error
			{
				ASSERT_COM_SUCCESS(hr, "Error while enumerating CA database schema columns");
			}
		}
		finally
		{
			CHECK_COM_NULL_AND_RELEASE(pEnumColumn);
			CHECK_COM_NULL_AND_RELEASE(pCertView);
		}

		_columns = retArray;
		return _columns;
	}

	bool isDispositionColumn(CertColumn ^column)
	{
		return "request.disposition" == column->name->ToLower();
	}

	bool isRequestIDColumn(CertColumn ^column)
	{
		return "requestid" == column->name->ToLower();
	}
	
	System::Data::DataTable ^CertificateServices::queryIssuedCertificates(System::Collections::Generic::ICollection<CertColumn ^> ^columnsReturn)
	{
		CertColumn ^colDisposition = System::Linq::Enumerable::Single<CertColumn ^>(columns, gcnew System::Func<CertColumn^,bool>(isDispositionColumn));

		return queryCertificates(
			gcnew SingleValueRestriction(RequestDisposition::Issued, colDisposition), 
			columnsReturn
		);
	}

	System::Data::DataTable ^CertificateServices::queryAllCertificates(System::Collections::Generic::ICollection<CertColumn ^> ^columnsReturn)
	{
		return queryCertificates(gcnew NoRestriction(), columnsReturn);
	}

	System::Data::DataRow ^CertificateServices::findCertificate(int requestID, System::Collections::Generic::ICollection<CertColumn ^> ^columnsReturn)
	{
		CertColumn ^colRequestID = System::Linq::Enumerable::Single<CertColumn ^>(columns, gcnew System::Func<CertColumn^,bool>(isRequestIDColumn));

		return queryCertificates(
			gcnew SingleValueRestriction(requestID, colRequestID), 
			columnsReturn
		)->Rows[0];

		//HRESULT hr;
		//ICertView2 *pCertView = NULL;

		//try
		//{
		//	pCertView = openCAConnection();

		//		// filter for the request number
		//	CertColumn ^colRequestID = nullptr;
		//	for each(CertColumn ^col in columns)
		//		if(isRequestIDColumn(col))
		//		{
		//			colRequestID = col;
		//			break;
		//		}
		//	if (nullptr == colRequestID)
		//		throw gcnew System::InvalidOperationException("CA does not contain a RequestID column!");

		//	VARIANT varRequestID;
		//	System::Runtime::InteropServices::Marshal::GetNativeVariantForObject(requestID, System::IntPtr(&varRequestID));
		//	hr = pCertView->SetRestriction(colRequestID->index, CVR_SEEK_EQ, CVR_SORT_NONE, &varRequestID);
		//	VariantClear(&varRequestID);
		//	ASSERT_COM_SUCCESS(hr, "Could not restrict certificate view to one specific certificate request ID");

		//	return queryCertificates(pCertView, columnsReturn)->Rows[0];
		//}
		//finally
		//{
		//	CHECK_COM_NULL_AND_RELEASE(pCertView);
		//}
	}

	//System::Data::DataTable ^CertificateServices::queryCertificates(int iDisposition, System::Collections::Generic::ICollection<CertColumn ^> ^columnsReturn)
	//{
	//	HRESULT hr;
	//	ICertView2 *pCertView = NULL;

	//	try
	//	{
	//		pCertView = openCAConnection();

	//		if (-1 != iDisposition)
	//		{
	//				// filter for the disposition given by iDisposition
	//			CertColumn ^colDisposition = nullptr;
	//			for each(CertColumn ^col in columns)
	//				if(isDispositionColumn(col))
	//				{
	//					colDisposition = col;
	//					break;
	//				}
	//			if (nullptr == colDisposition)
	//				throw gcnew System::InvalidOperationException("CA does not contain a disposition column, which is required to determine which of the certificates are valid");

	//			VARIANT varDisposition;
	//			System::Runtime::InteropServices::Marshal::GetNativeVariantForObject(iDisposition, System::IntPtr(&varDisposition));
	//			hr = pCertView->SetRestriction(colDisposition->index, CVR_SEEK_EQ, CVR_SORT_NONE, &varDisposition);
	//			VariantClear(&varDisposition);
	//			ASSERT_COM_SUCCESS(hr, "Could not restrict certificate view to issued (or other specific) certificates");
	//		}

	//		return queryCertificates(pCertView, columnsReturn);
	//	}
	//	finally
	//	{
	//		CHECK_COM_NULL_AND_RELEASE(pCertView);
	//	}
	//}

	System::Data::DataTable ^CertificateServices::queryCertificates(CertQueryRestriction ^filter, System::Collections::Generic::ICollection<CertColumn ^> ^columnsReturn)
	{
		ICertView2 *pCertView = NULL;

		try
		{
			pCertView = openCAConnection();
			filter->restrictCertView(pCertView);

			return queryCertificates(pCertView, columnsReturn);
		}
		finally
		{
			CHECK_COM_NULL_AND_RELEASE(pCertView);
		}
	}

	System::Data::DataTable ^CertificateServices::queryCertificates(ICertView2 *pCertView, System::Collections::Generic::ICollection<CertColumn ^> ^columnsReturn)
	{
		HRESULT hr;
		IEnumCERTVIEWROW *pEnumRow = NULL;
		
			// Prepare columns in returned DataTable
		System::Collections::Generic::IDictionary<int, CertColumn^> ^dictColumns
			= gcnew System::Collections::Generic::Dictionary<int, CertColumn^>;
		System::Data::DataTable ^dtRet = gcnew System::Data::DataTable();
		for each(CertColumn ^column in columnsReturn)
		{
			dtRet->Columns->Add(column->name);
			dictColumns->Add(column->index, column);
		}

		try
		{
				// Filter for result columns
			hr = pCertView->SetResultColumnCount(columnsReturn->Count);
			ASSERT_COM_SUCCESS(hr, "Could not set resulting column count for CA database query");

			for each (CertColumn ^column in columnsReturn)
			{
				hr = pCertView->SetResultColumn(column->index);
				ASSERT_COM_SUCCESS(hr, "test"); //System::String::Concat("Could not set resulting column ", column->name, " for CA database query"));
			}

			hr = pCertView->OpenView(&pEnumRow);
			ASSERT_COM_SUCCESS(hr, "Could not open view to enumerate certificates");

				// First find out the types of the columns
			hr = pEnumRow->Reset();
			ASSERT_COM_SUCCESS(hr, "Error when resetting row enumeration");

			IEnumCERTVIEWCOLUMN *pEnumColumns4Types = NULL;
			try
			{
				LONG lngCurrentColumn = 0, dummy=0;

				hr = pEnumRow->Next(&dummy);
				ASSERT_COM_SUCCESS(hr, "Could not access first row of certificate data");
				if (S_FALSE == hr)
					return nullptr;	// No certificates match the filters

				hr = pEnumRow->EnumCertViewColumn(&pEnumColumns4Types);
				ASSERT_COM_SUCCESS(hr, "Error when enumerating column for types");

				hr = pEnumColumns4Types->Reset();
				ASSERT_COM_SUCCESS(hr, "Could not reset column type enumeration");
				while (S_OK == (hr = pEnumColumns4Types->Next(&lngCurrentColumn)))
				{
					LONG iType;
					hr = pEnumColumns4Types->GetType(&iType);
					ASSERT_COM_SUCCESS(hr, "Could not retrieve type for column #" + lngCurrentColumn.ToString());
					switch (iType)
					{
					case PROPTYPE_DATE:
						dtRet->Columns[lngCurrentColumn]->DataType = System::DateTime::typeid;
						break;
					case PROPTYPE_LONG:
						dtRet->Columns[lngCurrentColumn]->DataType = long::typeid;
						break;
					case PROPTYPE_BINARY:
						//dtRet->Columns[lngCurrentColumn]->DataType = array<System::Byte>::typeid;
						//break;
					case PROPTYPE_STRING:
						dtRet->Columns[lngCurrentColumn]->DataType = System::String::typeid;
						break;
					}
				}
			}
			finally
			{
				CHECK_COM_NULL_AND_RELEASE(pEnumColumns4Types);
			}

			hr = pEnumRow->Reset();
			ASSERT_COM_SUCCESS(hr, "Error when resetting row enumeration");

			LONG lngCurrentRow = 0;
			while(S_OK == (hr = pEnumRow->Next(&lngCurrentRow)))	// writes the current index to currentIndex, not otherwise
			{	// Now we have one certificate, enumerate its attributes
				IEnumCERTVIEWCOLUMN *pEnumColumn = NULL;

				try
				{
					hr = pEnumRow->EnumCertViewColumn(&pEnumColumn);
					ASSERT_COM_SUCCESS(hr, "Error when enumerating row values");

					hr = pEnumColumn->Reset();
					ASSERT_COM_SUCCESS(hr, "Could not reset column value enumeration");

						// the certificate info will be written to drNew
					System::Data::DataRow ^drNew = dtRet->NewRow();

					LONG lngCurrentColumn = 0;
					while (S_OK == (hr = pEnumColumn->Next(&lngCurrentColumn)))
					{
						LONG iType;
						hr = pEnumColumn->GetType(&iType);
						ASSERT_COM_SUCCESS(hr, "Could not retrieve type for column #" + lngCurrentColumn.ToString());

						LONG lngOutFormatting;
						if (PROPTYPE_BINARY == iType)
							lngOutFormatting = CV_OUT_BASE64HEADER;
						else
							lngOutFormatting = CV_OUT_BINARY;

						VARIANT varColumnData;
						hr = pEnumColumn->GetValue(lngOutFormatting, &varColumnData);
						ASSERT_COM_SUCCESS(hr, System::String::Concat("Could not retrieve value for row #", lngCurrentRow.ToString(), ", column #", lngCurrentColumn.ToString()) );

						//if (PROPTYPE_BINARY == iType)
						//	drNew[lngCurrentColumn] = System::Convert::FromBase64String(
						//		System::Runtime::InteropServices::Marshal::GetObjectForNativeVariant(System::IntPtr(&varColumnData))->ToString()
						//	);
						//else
						//{
							System::Object ^val = System::Runtime::InteropServices::Marshal::GetObjectForNativeVariant(System::IntPtr(&varColumnData));
						
							if (val == nullptr)
								val = System::DBNull::Value;
						
							drNew[lngCurrentColumn] = val;
						//}

						VariantClear(&varColumnData);
					}

					dtRet->Rows->Add(drNew);
				}
				finally
				{
					CHECK_COM_NULL_AND_RELEASE(pEnumColumn);
				}
			}

			if (S_FALSE != hr && S_OK != hr)		// S_FALSE = enumeration ends without error
			{
				ASSERT_COM_SUCCESS(hr, "Error while enumerating certificate rows or value columns");
			}
		}
		finally
		{
			CHECK_COM_NULL_AND_RELEASE(pEnumRow);
		}

		return dtRet;
	}

#ifdef MSCAWRAPPER_WIN32_FUNCTIONS
	void CertificateServices::stopService()
	{
		HRESULT hr = CertSrvServerControl(L"gkchannserver1.chann.local", CSCONTROL_SHUTDOWN, NULL, NULL);
		ASSERT_COM_SUCCESS(hr, "could not stop CA server");
	}
#endif MSCAWRAPPER_WIN32_FUNCTIONS

}
}
