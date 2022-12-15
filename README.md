# Yelp-API-Data-Collector
With the inputs of a business type and location, return data on 100 nearby businesses in an Excel workbook.


## Setup

Create and activate a virtual environment:

```sh
conda create -n yelpAPI-env python=3.8

conda activate yelpAPI-env
```

Install package dependencies:

```sh
pip install -r requirements.txt
```

## Configuration


[Obtain an API Key](https://www.yelp.com/developers/documentation/v3/authentication) from Yelp.

Then create a local ".env" file and provide the key like this:

```sh
# this is the ".env" file...

YELP_API_KEY="_________"

```



## Usage

Run the program:

```sh
python app.py
```
