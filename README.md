# Facebook-API-Data-Collector
With the inputs of a Facebook business page, return data on 100 posts in an Excel workbook.


## Setup

Create and activate a virtual environment:

```sh
conda create -n facebookAPI-env python=3.8

conda activate facebookAPI-env
```

Install package dependencies:

```sh
pip install -r requirements.txt
```

## Configuration


[Obtain an API Key](https://developers.facebook.com/tools/explorer/) from Facebook.

Then create a local ".env" file and provide the key like this:

```sh
# this is the ".env" file...

FACEBOOK_API_KEY="_________"

```


## Usage

Run the program:

```sh
python app.py
```
