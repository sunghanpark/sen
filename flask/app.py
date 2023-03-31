from flask import Flask, render_template, request
from textblob import TextBlob
import pandas as pd

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/', methods=['POST'])
def analyze():
    # Get the input sentence(s) from the form
    sentences = request.form['sentences']
    sentences_list = sentences.split('\r\n')

    # Create empty lists to store the sentiment analysis results
    polarity_list = []
    subjectivity_list = []

    # Loop through each sentence and analyze the emotions
    for sentence in sentences_list:
        # Create a TextBlob object
        text_blob = TextBlob(sentence)

        # Get the sentiment analysis result
        polarity, subjectivity = text_blob.sentiment

        # Append the result to the lists
        polarity_list.append(polarity)
        subjectivity_list.append(subjectivity)

    # Create a pandas DataFrame object to store the sentiment analysis results
    data = {'Sentence': sentences_list, 'Polarity': polarity_list, 'Subjectivity': subjectivity_list}
    df = pd.DataFrame(data)

    # Save the DataFrame object to an Excel file
    df.to_excel('sentiment_analysis_result.xlsx', index=False)

    # Render the results template and pass the DataFrame object as a parameter
    return render_template('results.html', data=df.to_html())

if __name__ == '__main__':
    app.run(debug=True)
