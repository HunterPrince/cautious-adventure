import os
import logging
from adobe.pdfservices.operation.auth.credentials import ServicePrincipalCredentials
from adobe.pdfservices.operation.pdfservices import PDFServices
from adobe.pdfservices.operation.pdfservices_media_type import PDFServicesMediaType
from adobe.pdfservices.operation.create_pdf.create_pdf_from_word_params import CreatePDFFromWordParams
from adobe.pdfservices.operation.create_pdf.create_pdf_job import CreatePDFJob
from adobe.pdfservices.operation.create_pdf.create_pdf_result import CreatePDFResult
from adobe.pdfservices.operation.exception.exceptions import ServiceApiException, ServiceUsageException, SdkException
from adobe.pdfservices.operation.io.stream_asset import StreamAsset
from adobe.pdfservices.operation.io.cloud_asset import CloudAsset

# Initialize the logger
logging.basicConfig(level=logging.INFO)

class CreatePDFFromDOCXWithOptions:
    def __init__(self):
        try:
            with open('./createPDFInput.docx', 'rb') as file:
                input_stream = file.read()

            # Initial setup, create credentials instance
            credentials = ServicePrincipalCredentials(
                client_id=os.getenv('PDF_SERVICES_CLIENT_ID'),
                client_secret=os.getenv('PDF_SERVICES_CLIENT_SECRET')
            )

            # Creates a PDF Services instance
            pdf_services = PDFServices(credentials=credentials)

            # Creates an asset(s) from source file(s) and upload
            input_asset = pdf_services.upload(input_stream=input_stream, mime_type=PDFServicesMediaType.DOCX)

            # Create parameters for the job
            create_pdf_params = CreatePDFFromWordParams(document_language=DocumentLanguage.EN_US)

            # Creates a new job instance
            create_pdf_job = CreatePDFJob(input_asset=input_asset, create_pdf_params=create_pdf_params)

            # Submit the job and gets the job result
            location = pdf_services.submit(create_pdf_job)
            pdf_services_response = pdf_services.get_job_result(location, CreatePDFResult)

            # Get content from the resulting asset(s)
            result_asset: CloudAsset = pdf_services_response.get_result().get_asset()
            stream_asset: StreamAsset = pdf_services.get_content(result_asset)

            # Ensure the output directory exists
            output_dir = 'output'
            os.makedirs(output_dir, exist_ok=True)
            output_file_path = os.path.join(output_dir, 'CreatePDFFromDOCXWithOptions.pdf')

            # Creates an output stream and copy stream asset's content to it
            with open(output_file_path, "wb") as file:
                file.write(stream_asset.get_input_stream())

        except (ServiceApiException, ServiceUsageException, SdkException) as e:
            logging.exception(f'Exception encountered while executing operation: {e}')

if __name__ == "__main__":
    CreatePDFFromDOCXWithOptions()