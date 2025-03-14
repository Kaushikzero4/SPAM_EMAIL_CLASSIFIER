import streamlit as st
import pickle
from sklearn.feature_extraction.text import CountVectorizer
from win32com.client import Dispatch

# Full paths to the files
model_path = r'E:\New folder\SPAM_EMAIL_CLASSIFIER\spam.pkl'
vectorizer_path = r'E:\New folder\SPAM_EMAIL_CLASSIFIER\vectorizer.pkl'

# Load the model and vectorizer from the .pkl files
model, cv = pickle.load(open(model_path, 'rb'))  # Assuming the model and vectorizer were saved as a tuple

def main():
    # Set up the title and description
    st.title("SPAM EMAIL CLASSIFIER")
    st.write("This app helps classify emails as Spam or Not Spam using machine learning models.")

    # Sidebar options for activities
    activities = ["Classification", "About"]
    choice = st.sidebar.selectbox("Select Activities", activities)

    if choice == "Classification":
        st.subheader("Classify Your Email Text")
        msg = st.text_input("Enter the text of your email")

        if st.button("Process"):
            if msg.strip() == "":
                st.warning("Please enter some text to classify.")
            else:
                data = [msg]
                # Transform the input data using the vectorizer
                vec = cv.transform(data).toarray()
                result = model.predict(vec)

                if result[0] == 0:
                    st.success("This is Not a Spam email")
                else:
                    st.error("This is a Spam email")

    elif choice == "About":
        st.subheader("About this App")
        st.write("""
            The Spam Email Classifier helps you identify if an email is spam or not spam.
            Simply enter the email text, and the app predicts using machine learning whether it is spam or ham.
            It uses advanced algorithms to analyze the content and give accurate results.
            It utilizes a pre-trained model along with a `CountVectorizer` to process the email text.
            It provides a quick way to manage and filter the  inbox!
        """)

if __name__ == "__main__":
    main()

