FROM python:3.10.5

RUN mkdir -p /usr/src/app/

WORKDIR /usr/src/app/

COPY . /usr/src/app/

RUN pip install appdirs==1.4.4
RUN pip install beautifulsoup4==4.12.2
RUN pip install bs4==0.0.1
RUN pip install certifi==2023.7.22
RUN pip install charset-normalizer==3.2.0
RUN pip install colorama==0.4.6
RUN pip install cssselect==1.2.0
RUN pip install ebaysdk==2.2.0
RUN pip install et-xmlfile==1.1.0
RUN pip install fake-useragent==1.1.3
RUN pip install idna==3.4
RUN pip install importlib-metadata==6.8.0
RUN pip install lxml==4.9.3
RUN pip install numpy==1.25.2
RUN pip install openpyxl==3.1.2
RUN pip install pandas==2.0.3
RUN pip install parse==1.19.1
RUN pip install pyee==8.2.2
RUN pip install pyppeteer==1.0.2
RUN pip install pyquery==2.0.0
RUN pip install python-dateutil==2.8.2
RUN pip install python-dotenv==1.0.0
RUN pip install pytz==2023.3
RUN pip install requests==2.31.0
RUN pip install requests-html==0.10.0
RUN pip install six==1.16.0
RUN pip install soupsieve==2.4.1
RUN pip install tqdm==4.65.0
RUN pip install tzdata==2023.3
RUN pip install urllib3==1.26.16
RUN pip install w3lib==2.1.1
RUN pip install websockets==10.4
RUN pip install zipp==3.16.2

EXPOSE 8080

CMD [ "python", "main.py"]