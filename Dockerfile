FROM buildpack-deps:jessie

RUN echo 'deb http://ftp.us.debian.org/debian unstable main' >> /etc/apt/sources.list \ 
 && apt-get update \
 && apt-get -y install libreoffice libreoffice-writer ure libreoffice-java-common libreoffice-core libreoffice-common  \
 && apt-get -y install openjdk-8-jre fonts-opensymbol hyphen-fr hyphen-de hyphen-en-us hyphen-it hyphen-ru fonts-dejavu fonts-dejavu-core fonts-dejavu-extra fonts-droid fonts-dustin fonts-f500 fonts-fanwood fonts-freefont-ttf fonts-liberation fonts-lmodern fonts-lyx fonts-sil-gentium fonts-texgyre fonts-tlwg-purisa  \
 && apt-get -y install libreoffice-script-provider-python uno-libs3 python3-uno \
 && apt-get -y install python3 python3-flask python3-pip
RUN pip3 install gevent

RUN mkdir -p /usr/app
WORKDIR /usr/app

COPY . /usr/app

CMD ["python3", "main.py"]