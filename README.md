# Marathi to English Excel Translator

This project provides various tools to translate text from Marathi to English within Excel files. It caters to different needs, whether you want to use cloud-based services like the Azure Translation API, a simple Python library like `translate` for quick translations, or a completely offline approach using machine learning models from the `transformers` library.

## Features

- **Azure Translation API**: Leverage the cloud-based Azure Translation API for high-quality translations.
- **Translate Library**: A simple wrapper around Google Translate for easy use.
- **Offline Translation**: Use pre-trained models from `transformers` for offline translations.
- **Batch Processing**: Processes each sheet in the workbook and translates each cell.
- **Progress Tracking**: Detailed print statements to monitor the translation progress.
- **Checkpointing**: Saves progress to avoid re-translating in case of interruptions.

## Requirements

- Python 3.6+
- `openpyxl` library for handling Excel files.
- `requests` library for making API requests (for Azure Translation API).
- `translate` library for using the `translate` Python wrapper.
- `transformers` and `sentencepiece` libraries for offline translation models.

## Installation

To install the required libraries, run:

```bash
pip install openpyxl requests translate transformers sentencepiece
```

## Usage

There are three main scripts in this project:

1. `translate_azure.py` - for translating using the Azure Translation API.
2. `translate_simple.py` - for translating using the `translate` library.
3. `translate_offline.py` - for translating using offline `transformers` models.

To run a script, use:

```bash
python <script_name>.py
```

Ensure you edit the script to set your input and output file names, and configure any API keys or model names as needed.

## Configuration

Each script has specific configuration options:

- Azure Translation API:
  - Set `SUBSCRIPTION_KEY`, `ENDPOINT`, and `LOCATION` with your Azure credentials.
  - Configure rate limiting and error handling as per your Azure plan.

- Translate Library:
  - This method is the simplest and requires no special configuration.
  - Please note that this still makes online API calls under the hood.

- Offline Transformers Model:
  - Choose and download a suitable pre-trained model from the Hugging Face Model Hub.
  - Ensure the model supports Marathi to English translation.

## Contributing

We welcome contributions to improve the tools provided in this project. If you have fixes, improvements, or additional translation methods, please fork the repository and submit a pull request with your changes.

## License

This project is open source and available under the [MIT License](LICENSE).

## Acknowledgements

- Azure Translation API is a product of Microsoft Azure Cognitive Services.
- The `translate` library is a simple interface for the Google Translate API.
- The offline translation feature uses models from the [Hugging Face Transformers](https://huggingface.co/transformers/) library.

## Disclaimer

The accuracy of translations may vary based on the method used. Cloud-based services often provide higher quality translations compared to offline methods. The `translate` library is a third-party tool and may be subject to API limitations and changes.

## Motivation

To help an organisation to make their process faster and earn money (INR 2000). 