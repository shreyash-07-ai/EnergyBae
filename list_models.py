import google.generativeai as genai

GEMINI_API_KEY = "AIzaSyCSp9ccFGxAPnV9zh_vFzAPhzu8SkugcoM"
genai.configure(api_key=GEMINI_API_KEY)

print("Available models:")
for model in genai.list_models():
    print(f"- {model.name}")
    if hasattr(model, 'supported_generation_methods'):
        print(f"  Supported methods: {model.supported_generation_methods}")
