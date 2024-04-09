from __future__ import absolute_import, division, print_function
import pathlib
import pandas as pd
import seaborn as sns
import tensorflow as tf
from tensorflow import keras
from tensorflow.keras import layers
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

def norm(x):
    # Regular normalisation does not work here as the std is 0 sometimes and is giving a NAN
    #return (x-train_stats['mean'])/train_stats['std']
    return(x/100)

def prep_training_data(self):
    # DP - create training data set (using just the first file for now, others to be added via concat later)
    self.raw_dataset = self.joined_stored_data
    self.dataset = self.raw_dataset.copy()  # make a copy to not tamper with original data

    #get rid of first name and surname column
    # DP - might need to check if they exist first, or just don't show them to user in the first place?
    # DP - may want to put this in a try statement, also the gradebook files have 'last name' but binesh code had 'surname'
    to_drop = ['First name', 'Last name', 'Surname']
    #self.dataset = self.dataset.drop(to_drop, axis=1)  # drop the items that are not assessed yet
   #to_drop = ['First name', 'Last name']
    self.dataset = self.dataset.drop(columns=[col for col in to_drop if col in self.dataset.columns])
    # DP - is this split for testing needed in the final code?
    #DP - comments below to use full data set for training or just 80% of it
    #self.train_dataset = self.dataset.sample(frac=0.8, random_state=0)  # split into 80% training and 20% testing
    self.train_dataset = self.dataset
    self.test_dataset = self.dataset.drop(self.train_dataset.index)


    self.train_labels = self.train_dataset.pop('Unit total')
    self.test_labels = self.test_dataset.pop('Unit total')

    # create normalised data set, by dividing all by 100
    # DP - how does this manage /zero, and what if the column is not out of 100?
    self.normed_train_data = norm(self.train_dataset)
    self.normed_test_data = norm(self.test_dataset)

def prep_predict_data(self):
    #replaced binesh's raw_dataset_2022 with self.raw_predict_dataset
    self.raw_predict_dataset = self.stored_data_1_p
    self.predict_dataset = self.raw_predict_dataset.copy()
    to_drop = ['First name', 'Last name', 'Surname', 'Unit total']
    #self.predict_dataset = self.predict_dataset.drop(to_drop, axis=1)
    #to_drop = ['First name', 'Last name', 'Surname']
    # self.dataset = self.dataset.drop(to_drop, axis=1)  # drop the items that are not assessed yet
    # to_drop = ['First name', 'Last name']
    self.predict_dataset = self.predict_dataset.drop(columns=[col for col in to_drop if col in self.predict_dataset.columns])

    # DP - Why note use norm function here?
    self.normed_predict_dataset = self.predict_dataset / 100

def predict_grades(self):
    # DP- does this need to be flattened? in some parts of BPV code it is, in other parts it's not
    #loss, mae, mse = self.model.evaluate(self.normed_test_data, self.test_labels, verbose=0)
    #print("Testing set Mean Abs Error: {:5.2f} marks".format(mae))

    self.test_predictions = self.model.predict(self.normed_predict_dataset).flatten()
    self.df_prediction = pd.DataFrame(self.test_predictions, columns=['Predicted Marks'])
    self.df_prediction['Student'] = self.raw_predict_dataset['First name']
    #enable this only if there is stored data existing
    #self.df_prediction['Historical_marks'] = self.raw_dataset['Unit total']

    self.df_Pred = pd.DataFrame()
    #enable this only if there is stored data( e.g. not a model load)
    #generate_marks_histogram(self)
    self.df_Pred['Student_FirstName'] = self.raw_predict_dataset['First name']
    if 'Last name' in self.raw_predict_dataset.columns:
        self.df_Pred['Student_LastName'] = self.raw_predict_dataset['Last name']
    elif 'Surname' in self.raw_predict_dataset.columns:
        self.df_Pred['Student_LastName'] = self.raw_predict_dataset['Surname']
    self.df_temp_Pred = []
    self.df_temp_Pred = pd.DataFrame(self.test_predictions, columns=['Header'])
    self.df_Pred['Predicted Marks'] = self.df_temp_Pred['Header']




def generate_marks_histogram(self):
    # below is the historgram of historical total marks
    sns.set(font_scale=1.3)
    sns_plot = sns.histplot(self.raw_dataset['Unit total'], binwidth=3, color='red', label='Historical', legend=True)

    # below is the histogram of predicted total marks
    # sns.set(font_scale=1.3)
    sns_plot = sns.histplot(self.df_prediction['Predicted Marks'], binwidth=3, label='Predicted', legend=True)

    fig = sns_plot.get_figure()
    fig.savefig("output.jpg")

def build_model(self):
    # DP -should this be using self.train_dataset or self.normed_train_data?
    self.model=keras.Sequential([
      layers.Dense(8, activation=tf.nn.relu, input_shape=[len(self.train_dataset.keys())]),
      layers.Dense(8,activation=tf.nn.relu),
      layers.Dense(1)
    ])
    optimizer=tf.keras.optimizers.RMSprop(0.001)

    self.model.compile(loss='mse',
                optimizer=optimizer,
                metrics=['mae','mse'])

def trial_run(self):
    #trial run to see if the model is working

    #restart progress bar
    self.train_progress_bar.set(0)
    # DP - should this use regular data set or normalised one? One is commented out in BPV code
    example_batch=self.train_dataset[:10]
    # example_batch=self.normed_train_data[:10]
    example_result=self.model.predict(example_batch)
    example_result

    EPOCHS = 1000

    history = self.model.fit(
        self.normed_train_data, self.train_labels,
        epochs=EPOCHS, validation_split=0.2, verbose=0,
        callbacks=[PrintDot(self.train_progress_bar)])
    hist = pd.DataFrame(history.history)
    hist['epoch'] = history.epoch
    hist.tail()

class PrintDot(keras.callbacks.Callback):

    def __init__(self, progress_bar):
        super().__init__()
        self.progress_bar = progress_bar
    def on_epoch_end(self,epoch, logs):
        if epoch % 100 ==0:
            print('')
        #if epoch % 10 == 0:
        progress = epoch / 1000
        self.progress_bar.set(progress)
        self.progress_bar.update_idletasks()

        print('.',end='')
