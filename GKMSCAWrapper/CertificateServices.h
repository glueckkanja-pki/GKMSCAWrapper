#pragma once

#include "MSCAObject.h"
#include "CertColumn.h"
#include "certqueryrestriction.h"

namespace GK {
namespace MSCAWrapper {

	/// <summary>
	/// This class provides methods to access Microsoft Certificate Services (MS CA). It wraps around
	/// CryptoAPI to access the Certificate Services database and its administrative functions (like
	/// revocation).
	/// </summary>
	public ref class CertificateServices : public MSCAObject
	{
	public:
		/// <summary>
		/// Establish a new handle to Microsoft Certificate Services
		/// </summary>
		/// <param name="caConfiguration">A CA configuration string of the form "DNSName/CAName" as
		/// specified the MSDN library for ICertView2::OpenConnection.</param>
		CertificateServices(System::String ^caConfiguration);

		enum class RevocationReason {
			UNSPECIFIED = 0, 
			KEY_COMPROMISE = 1, 
			CA_COMPROMISE = 2, 
			AFFILIATION_CHANGED = 3, 
			SUPERSEDED = 4, 
			CESSATION_OF_OPERATION = 5, 
			CERTIFICATE_HOLD = 6 
		};

		enum class RequestDisposition {
			Issued = 20,
			Revoked = 21
		};

		/// <summary>
        /// Revokes a certificate in the MS CA database.
        /// </summary>
        /// <param name="strSerial">The serial number of the certificate to be revoked. This string
		/// must be an even number of hexadecimal digits representing the binary value of the
		/// serial. See MSDN documentation for ICertAdmin2::RevokeCertificate for details.</param>
		/// <param name="reason">Why revoke this certificate (will be written to CRL)?</param>
		/// <param name="date">At which time will the certificate become invalid (may be in the past)</param>
		void revokeCertificate(System::String ^strSerial, RevocationReason reason, System::DateTime date);

		/// <summary>
		/// Deletes a row in the MS CA database within the requests table.
		/// </summary>
		/// <param name="requestID">The Request ID (sometimes also referred to as database ID)
		/// of the certificate to delete from the database.</param>
		void deleteCertificateRow(int requestID);

		/// <summary>
		/// Imports a certificate into the MS CA database. This certificate must have been published
		/// by the CA certificate of the MS CA, the signature will be checked.
		/// </summary>
		/// <param name="baCertificate">DER encoded X.509 certificate</param>
		/// <returns>The request ID of the imported certificate</returns>
		int importCertificate(array<System::Byte> ^baCertificate);

		/// <summary>
		/// Imports a certificate into the MS CA database. This certificate must have been published
		/// by the CA certificate of the MS CA, the signature will be checked.
		/// </summary>
		/// <param name="sCertificate">PEM-encoded X.509 certificate</param>
		/// <returns>The request ID of the imported certificate</returns>
		int importCertificate(System::String ^sCertificate);

		/// <summary>
		/// Gets a list of the columns available in the database table of all 
		/// certificates and certificate requests
		/// </summary>
		property System::Collections::Generic::IList<CertColumn ^> ^columns{
			System::Collections::Generic::IList<CertColumn ^> ^get();
		}

		/// <summary>
		/// Fetches all certificates issued by the certification authority. Revoked
		/// certificates and pending requests are excluded, but expired certificates are returned. (Disposition is 20)
		/// </summary>
		/// <param name="columnsReturn">The columns that should be queried for each certificate.
		/// A list of all availabe columns can be acquired via the columns property.</param>
		/// <returns>A DataTable with the requested columns, one row represents one certificate.</returns>
		System::Data::DataTable ^queryIssuedCertificates(System::Collections::Generic::ICollection<CertColumn ^> ^columnsReturn);

		/// <summary>
		/// Fetches all certificates (valid, expired, revoked, requested, failed) in the certification authority database.
		/// </summary>
		/// <param name="columnsReturn">The columns that should be queried for each certificate.
		/// A list of all availabe columns can be acquired via the columns property.</param>
		/// <returns>A DataTable with the requested columns, one row represents one certificate.</returns>
		System::Data::DataTable ^queryAllCertificates(System::Collections::Generic::ICollection<CertColumn ^> ^columnsReturn);

		System::Data::DataRow ^findCertificate(int requestID, System::Collections::Generic::ICollection<CertColumn ^> ^columnsReturn);

		/// <summary>
		/// Fetches all certificates that match the given filter
		/// </summary>
		/// <param name="filter">Filter for the certificates being returned.</param>
		/// <param name="columnsReturn">The columns that should be queried for each certificate.
		/// A list of all availabe columns can be acquired via the columns property.</param>
		/// <returns>A DataTable with the requested columns, one row represents one certificate.</returns>
		System::Data::DataTable ^queryCertificates(CertQueryRestriction ^filter, System::Collections::Generic::ICollection<CertColumn ^> ^columnsReturn);

#ifdef MSCAWRAPPER_WIN32_FUNCTIONS
		/// <summary>
		/// Stops the CA service. No other methods can be called on the stopped service and this class
		/// does not provide a programmatical way to restart the service afterwards.
		/// </summary>
		void stopService();
#endif MSCAWRAPPER_WIN32_FUNCTIONS

	protected:
		
		/// <summary>
		/// Fetches certificates from the certificate services database.
		/// Restrictions and filters may have been applied before.
		/// </summary>
		/// <param name="pCertView">The connection to certificate services must have already been opened.
		/// Restrictions on this ICertView2 may have already been set to filter the output.</param>
		/// <param name="columnsReturn">The columns that should be queried for each certificate.
		/// A list of all availabe columns can be acquired via the columns property.</param>
		/// <returns>A DataTable with the requested columns, one row represents one certificate.</returns>
		System::Data::DataTable ^queryCertificates(ICertView2 *pCertView, System::Collections::Generic::ICollection<CertColumn ^> ^columnsReturn);
	private:
		System::Collections::Generic::IList<CertColumn ^> ^_columns;
	};
}
}