FROM python:2.7

RUN pip install openpyxl
RUN pip install email
RUN pip install Jinja2

RUN mkdir /app
WORKDIR /app

COPY . /app

RUN ["python"]
