# coding: utf8

import win32com.client

from sklearn.ensemble import RandomForestClassifier
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.externals import joblib

from os import listdir
from os.path import isfile, join
import os

from datetime import datetime,timedelta
import pytz
from timeit import default_timer as timer

from bs4 import BeautifulSoup  
import nltk
from nltk.corpus import stopwords # Import the stop word list
import re

import sys
import pdb


#start global definitions
X=[]
y=[]
outlook_folders=[]
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
path_file_vectorizer='data\\vectorizer.pkl'
path_file_classifier='data\\clf.pkl'
intelligent_folder_identifier="."
nb_clf_estimators=1000 #Number of estimaters used in our model (random forest classifier)
nb_max_minutes_to_classify_mail=-1 #We will each classify each mail older than less than nb_max_minutes_to_classify_mail in each folder
#end global definitions


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
			

# ============================================
# =========== main functions =================
# ============================================

def main_train_model_from_folders():
	print("Start training")
	start = timer()
	loop_through_folder(0,inbox.Folders,set_dataset)
	[vectorizer,clf,score]=train_model(nb_clf_estimators)
	end = timer()

	print("Score=",score,"%")
	print("Total time taken to train the model:",end-start,"s")

	#Dumps vectorizer and classifier (random forest classifier) in files such as to be retrieve later when we want to classify a mail
	joblib.dump(vectorizer, path_file_vectorizer)
	joblib.dump(clf, path_file_classifier)

	print("End training")


def main_predict_category_for_each_mail():
	print("Start prediction")
	#Loads vectorizer and classifier previously stored after having run the main_fetch_mails_and_train_model function
	start = timer()
	vectorizer=joblib.load(path_file_vectorizer) 
	clf=joblib.load(path_file_classifier) 

	loop_through_folder(0,inbox.Folders,lambda folder,k: outlook_folders.append(folder))

	predict_mail_category(vectorizer,clf,outlook_folders,inbox)
	predict_mail_category(vectorizer,clf,outlook_folders,inbox.Folders["Not Classified"])
	end = timer()

	print("End prediction")
	print("Total time taken to classify the mails:",end-start,"s")


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