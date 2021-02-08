#!/bin/bash

docker image list | grep pyinstaller_img:latest && docker rmi pyinstaller_img:latest