import xlwings as xw
import openai
openai.api_key = "xxxxxxxxxx"
model_engine = "text-davinci-003"

def main():
    wb=xw.Book.caller()
    wb.sheets[0].range("A1").value="測試"

def hello(name):
    return "hello {0}".format(name)

def main01():
    wb=xw.Book.caller()
    prompt = wb.sheets[0].range("A1").value
    completion = openai.Completion.create(
    engine=model_engine,
    prompt=prompt,
    max_tokens=128,
    n=1,
    stop=None,
    temperature=0.5,)
    response = completion.choices[0].text
    wb.sheets[0].range("A3").value=response    
    #print(response)
if __name__== "__main__":
    xw.books.active.set_mock_caller()
    main()
