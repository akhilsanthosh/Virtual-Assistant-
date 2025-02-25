import json
import pickle
from tensorflow.keras.models import load_model # type: ignore

with open('intents.json') as file:
    data = json.load(file)

model  = load_model('chat_model.h5')

# Saving tokenizer
with open("tokenizer.pkl", "rb") as f:
    tokenizer = pickle.load(f)
