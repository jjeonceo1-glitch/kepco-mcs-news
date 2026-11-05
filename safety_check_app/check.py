import google.generativeai as genai
# 실제 사용 중인 API 키로 변경하세요
genai.configure(api_key="AIzaSyCuREucMTocsNAyUXBLVQlKVdlrGoLnx9o") 

for m in genai.list_models():
    if 'generateContent' in m.supported_generation_methods:
        print(m.name)