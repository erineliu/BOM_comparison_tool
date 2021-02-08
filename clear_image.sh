#!/bin/bash

docker image list | grep -E "pyinstaller_img.*latest" && docker rmi pyinstaller_img:latest