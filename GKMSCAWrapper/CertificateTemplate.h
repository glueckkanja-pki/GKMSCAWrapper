#pragma once

namespace GK {
	namespace MSCAWrapper {
		/// <summary>
		/// This class provides access to the Certificate Templates configured in AD and in each CA.
		/// It makes heavy use of the ScrdEnrl.dll COM DLL found on Windows Servers 
		/// (and pre-Vista Windows OSes, but these old versions are not compatible with the code here)
		/// </summary>
		public ref class CertificateTemplate
		{
		public:
			static array<CertificateTemplate ^> ^RetrieveAllUserCertificateTemplates();

			const System::String ^TemplateName;

			const System::String ^TemplateOID;

			enum class TemplateCategory {
				USER = 1,
				MACHINE = 2
			};

			const TemplateCategory TemplateType;
		private:
			CertificateTemplate(System::String ^templateName, System::String ^templateOID, TemplateCategory templateCategory);
		};

	}
}