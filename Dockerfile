## set base image ##
FROM centos/python-36-centos7:latest

USER root

## set env var ##
ENV work_dir=/BOM_comparison_tool

## set work folder, if the folder isn't exist, it will auto create the new one ##
WORKDIR  $work_dir

ENV virtual_env=$work_dir/venv01


## create venv environment ##
RUN python3 -m venv --without-pip $virtual_env


## copy current files to WebTester ##
ADD . $work_dir
#ADD requirements.txt $work_dir


ENV PATH="$virtual_env/bin:$PATH"


## install gcc libev-dev ##
# RUN apt-get update
# RUN apt-get -y install gcc
# RUN apt-get -y install libev-dev
# RUN apt-get install -y iputils-ping


## upgrade pip ##
RUN pip3 install --upgrade pip

## install python package ##
RUN pip3 install -r requirements.txt

#CMD ["bash"]