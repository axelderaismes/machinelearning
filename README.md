
# Machine Learning with Microsoft Outlook

Do you spend a lot of time sorting and filtering your mails in Outlook?
Have you enough to create filtering rules for each new type of mail and modify them each time a mail is a little different that what you specified?
If yes, this tutorial is for you!

I had the same problem, receiving several hundreds to thousands of mail each day, and creating endless rules that slowed down my outlook a lot. I then developed a script in Python using sklearn (a machine learning library for Python) to automatically sort my mails.

Indeed, machine learning applies perfectly in this case, because :
- I have already a lot of mails (lots of data)
- These mails are already well sorted
- They don't follow extremely simple rules (such as "move all the mails containing 'order' in the folder 'Order') but rather more complex rules that will take some times to define and which are prone to be modified with future mails (such as "move all the mails containing "order" from this person to this other person with these other keyword A,B,...,Z and without this keyword 1,2,... from this date to this date..)


The code I will introduce below contains two main parts :
- The trainer : part that will train the model using all the folders specified (*) with the emails inside
- The predictor : part that will predict the category of a new received mail

I use in the model a random forest (RF) classifier which, after experimentations, is in this case the fastest and efficient model to classify the mails.
For information, I tried to use neural networks (MLP with Keras) for this task but it was much more time and CPU consumming for no or very little gain. Better use the best tools for a specific task !


NB : In this project, You will find a Jupyter Notebook that you can run with :
```bash
jupyter notebook
```
or you can the python file like this :
- To train :
```bash
python outlook_train_and_predict.py train
```
- To predict:
```bash
python outlook_train_and_predict.py predict
```



```python
# coding: utf8

import win32com.client #win32 allows to access microsoft applications like outlook

from sklearn.ensemble import RandomForestClassifier #our model with use a random forest
from sklearn.feature_extraction.text import CountVectorizer #we will need to normalise our input data before feeding it to our model
from sklearn.externals import joblib #we will need to save our model after the training to retrieve it before the prediction

from os import listdir
from os.path import isfile, join
import os

from datetime import datetime,timedelta
import pytz
from timeit import default_timer as timer

from bs4 import BeautifulSoup #contains a method to extract raw text from html data
from nltk.corpus import stopwords # Import the stop word list
import re #regular expressions

import sys
import pdb
```
Note that above we import nltk.corpus
If you haven't already installed the Natural Language Toolkit or one of the corpus needed, you will have to do the following :
- pip install nltk
- python
- import nltk
- nltk.download()
- Select "Corpora/Stopwords" and download
Below we define the global definitions :
X is the input vector
y is output vector
outlook_folders will list all the folders containing the intelligent_folder_identifier (by default, the folders containing '.'
path_file_vectorizer and path_file_classifier are the paths where we will save the vectorize and the classifier
nb_clf_estimators is the number used for the random forest (1000 is a good value, not too big, not too small)
nb_max_minutes_to_classify_mail is the number maximum of ancienty for a mail used when we retrieve the mails for predictions, -1 to classify all the mails not already classified


```python
#start global definitions
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) #6= Inbox

X=[]
y=[]
outlook_folders=[]

#you might want to modify the following variables
path_file_vectorizer='data//vectorizer.pkl'
path_file_classifier='data//clf.pkl'
intelligent_folder_identifier="." #We define "." as the identifier to identify the 'smart' folders
nb_clf_estimators=1000 #Number of estimators used in our model (random forest classifier)
nb_max_minutes_to_classify_mail=-1 #We will each classify each mail older than less than nb_max_minutes_to_classify_mail in each folder
#end global definitions
```

The function 'get_relevant_info_from_mail' keeps only essential informations about the email, here :
the sender name, the sender email address, the recipient, the message, the CC, the BCC, the subject and the body.
I don't store the dates but that would be interesting if a certain order in the mail will change the category. In this case a LSTM neural network might be more efficient...


```python
def get_relevant_info_from_mail(msg):
	#This function keeps only the relevant info from the mail
	try:
		return msg.SenderName+" "+msg.SenderEmailAddress+msg.To+msg.CC+msg.BCC+msg.Subject+msg.Body
	except Exception as e:
		print("Error in get_relevant_info_from_mail")
		try:
			print('Error on line {}'.format(sys.exc_info()[-1].tb_lineno), type(e).__name__, e)
		except:
			pass
```

The function 'strip_text' will get the output of 'get_relevant_info_from_mail' and strips it from all undesirable data like html tags, non-letters characters, remove the stop words and finally get a cleaned string with only real words


```python
def strip_text(raw_text):
	# This function strips raw text from all html tags, accentuated characters, remove stop words and keep only meaningful words
    # 1. Remove HTML
    review_text = BeautifulSoup(raw_text,"lxml").get_text() 
    # 2. Remove non-letters       
    accentedCharacters = "àèìòùÀÈÌÒÙáéíóúýÁÉÍÓÚÝâêîôûÂÊÎÔÛãñõÃÑÕäëïöüÿÄËÏÖÜŸçÇßØøÅåÆæœ" 
    letters_only = re.sub("[^a-zA-Z"+accentedCharacters+"\.]", " ", review_text) 
    # 3. Convert to lower case, split into individual words
    words = letters_only.lower().split()
    # 4. In Python, searching a set is much faster than searching a list, so convert the stop words to a set
    stops = set(stopwords.words("english"))
    # 5. Remove stop words
    meaningful_words = [w for w in words if not w in stops]
    # 6. Join the words back into one string separated by space, and return the result.
    return( " ".join( meaningful_words ))
```

The function train_model is the most interesting part.
X is the input dataset. Each input of X is a string containing only real words which is the output of strip_text.
But for a random forest, it is not exploitable as such.
You will need to convert each of these strings in an array of 0s and 1s.
The method to do that is to use the CountVectorizer.

In summary, the CountVectorizer will find all the words used in each of the strings of X and create a global dictionnary with these words. Each string of X wil then be compared to the global dictionnary and convert to an array of 0s an 1s such as this example :
Consider a very simple dictionnary : ["Machine","Learning","is","very","great","awesome","good"]
and a string 1: "Machine Learning is awesome"
and a string 2: "Learning is great"
You will then have the string 1 the following array : [1,1,1,0,0,1,0]
and for the string 2 : [0,1,1,0,1,0,0]

Two more things to be noted about CountVectorizer :
max_df = 0.95
min_df = 0.005

These two parameters allows us to reduce the dimensionality of our X matrix by using TF-IDF to identify un-important words. 

The min_df paramter makes sure we exclude words that only occur very rarely
The default also is to exclude any words that occur in every string (like 'the' ,'is', 'are'...)

We are excluding all words that occur in too many or too few documents, as these are very unlikely to be discriminative. Words that only occur in one document most probably are names, and words that occur in nearly all documents are probably stop words.


Now that we have a normalized X (X1) and y, which is our output dataset containing the identifiant of the category (0 for the first category, 1 for the second category, 2 for the third category...), we can train the model

The function returns the vectorizer, the classifier and the score. The latter is not useful in the training but can be used for benchmarking or following the score over time


```python
def train_model(nb_clf_estimators):
	# The min_df paramter makes sure we exclude words that only occur very rarely
	# The default also is to exclude any words that occur in every movie description
	# We are excluding all words that occur in too many or too few documents, as these are very unlikely to be discriminative
	# More about these parameters on : https://spandan-madan.github.io/
	vectorizer = CountVectorizer(analyzer = "word", tokenizer = None, preprocessor = None, stop_words = None, max_features = 5000,max_df=0.95, min_df=0.005)
	X1=vectorizer.fit_transform(X)

	# Here we train train a random forest because it is quick to train and works very well for text classification
	clf = RandomForestClassifier(n_estimators = nb_clf_estimators) 
	clf = clf.fit(X1, y)
	score=clf.score(X1, y)

	return [vectorizer,clf,score]
```

The function 'set_dataset' retrieves the relevant informations from each mail from each folders passed as parameters and adds for each mail a new line in X (relevant string) and y (category).
Note that we only deal with message with class=43 because events, notes don't have necessarily the same fields that the standard messages and need a specific treatment (not implemented in this tutorial).


```python
def set_dataset(folder,value):
	for msg in reversed(folder.Items):
		try:
			if msg.Class==43: #We only want to keep pure mails (this exclude appointments and notes for example)
				txt0=get_relevant_info_from_mail(msg)
				if txt0!=None:
					txt=strip_text(txt0)
					X.append(txt)
					y.append(value)
		except Exception as e:
			print("error when adding this mail to our dataset",msg.Subject)
			try:
				print('Error on line {}'.format(sys.exc_info()[-1].tb_lineno), type(e).__name__, e)
			except:
				pass
```

The function 'predict_mail_category' is used to predict the category and move accordingly all the mails inside the folder passed as parameters.
The function takes into input the vectorizer and the classifier and the folders (array that contains all the smart folders).


```python
def predict_mail_category(vectorizer,clf,folders,folder):
	dt=pytz.utc.localize(datetime.utcnow())+timedelta(minutes=-nb_max_minutes_to_classify_mail)

	for msg in reversed(folder.Items):
		try:
			if msg.Class==43:
				if msg.sentOn>dt or nb_max_minutes_to_classify_mail==-1:
					txt0=get_relevant_info_from_mail(msg)
					if txt0!=None:
						txt=strip_text(txt0)
						X1=vectorizer.transform([txt])
						prediction=clf.predict(X1)
						msg.Move(folders[prediction[0]])
		except Exception as e:
			print("error when adding this mail to our dataset",msg.Subject)
			try:
				print('Error on line {}'.format(sys.exc_info()[-1].tb_lineno), type(e).__name__, e)
			except:
				pass
```

The function 'loop_through_folder' iterates recursively through all the smart folders and smart subfolders in the inbox and lauching the function action passed as parameters


```python
def loop_through_folder(k,folders,action):
#Recursive function that iterate through folder and subfolder and launch the action method if there is a "." in the folder name
	for folder in folders:
		if intelligent_folder_identifier in folder.Name:
			# print("We will add the mails of this folder in our dataset",folder.Name,"with y=",k)
			action(folder,k)
			k=k+1

		try:
			l=len(folder.Folders)
			loop_through_folder(k,folder.Folders,action)
		except:
			pass
```

'main_train_model_from_folders' is the main function to train the random forest.
It will display the total score and time taken to train the model
and then will save the vectorizer and the classifier for future use.


```python
def main_train_model_from_folders():
	print("Start training")
	start = timer()
	loop_through_folder(0,inbox.Folders,set_dataset)
	[vectorizer,clf,score]=train_model(nb_clf_estimators)
	end = timer()

	print("Score=",score*100,"%")
	print("Total time taken to train the model:",end-start,"s")

	#Dumps vectorizer and classifier (random forest classifier) in files such as to be retrieve later when we want to classify a mail
	joblib.dump(vectorizer, path_file_vectorizer)
	joblib.dump(clf, path_file_classifier)

	print("End training")
```

'main_predict_category_for_each_mail' is the function used to move all the mails according the predicted categories.


```python
def main_predict_category_for_each_mail():
	print("Start prediction")
	#Loads vectorizer and classifier previously stored after having run the main_fetch_mails_and_train_model function
	start = timer()
	vectorizer=joblib.load(path_file_vectorizer) 
	clf=joblib.load(path_file_classifier) 

	loop_through_folder(0,inbox.Folders,lambda folder,k: outlook_folders.append(folder))

	predict_mail_category(vectorizer,clf,outlook_folders,inbox)
	end = timer()

	print("End prediction")
	print("Total time taken to classify the mails:",end-start,"s")
```

When executing the program ouside jupyter, you can specify if you want to train or predict.
To do so :
- To train : python outlook_train_and_predict.py train
- To predict : python outlook_train_and_predict.py predict


```python
if __name__ == '__main__':
	if len(sys.argv)!=2:
		action=""
	else:
		action=sys.argv[1]

	if action=="train":
		main_train_model_from_folders()
	elif action=="predict":
		main_predict_category_for_each_mail()
	else:
		print("Unknown keyword, use 'train' to train the model and 'predict' to predict the category of each mail")
```

    Unknown keyword, use 'train' to train the model and 'predict' to predict the category of each mail
    


```python
main_train_model_from_folders()
```

    Start training
    Score= 100.0 %
    Total time taken to train the model: 90.29 s
    End training
    


```python
main_predict_category_for_each_mail()
```

    Start prediction
    End prediction
    Total time taken to classify the mails: 3.038297298569887 s
    

# Author

Axel de Raismes

- https://www.linkedin.com/in/axelderaismes/
- http://www.thegeeklegacy.com
- https://twitter.com/axelderaismes