import google.generativeai as genai

GEMINI_API_KEY = "AIzaSyDWEtp2aEzQ76kNBT16d8AXBUmcplNKoXA"
genai.configure(api_key=GEMINI_API_KEY)

print("Available models:")
for model in genai.list_models():
    print(f"- {model.name}")
    if hasattr(model, 'supported_generation_methods'):
        print(f"  Supported methods: {model.supported_generation_methods}")
