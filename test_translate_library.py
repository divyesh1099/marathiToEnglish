from translate import Translator

# Assuming you want to translate from Marathi to English
translator = Translator(from_lang="mr", to_lang="en")

translation = translator.translate("आपले स्वागत आहे")
print(translation)