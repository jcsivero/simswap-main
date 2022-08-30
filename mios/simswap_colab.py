# -*- coding: utf-8 -*-
"""SimSwap colab.ipynb

Automatically generated by Colaboratory.

Original file is located at
    https://colab.research.google.com/github/neuralchen/SimSwap/blob/main/SimSwap%20colab.ipynb

This is a simple example of SimSwap on processing video with multiple faces. You can change the codes for inference based on our other scripts for image or single face swapping.

Code path: https://github.com/neuralchen/SimSwap

Paper path: https://arxiv.org/pdf/2106.06340v1.pdf or https://dl.acm.org/doi/10.1145/3394171.3413630
"""

## make sure you are using a runtime with GPU
## you can check at Runtime/Change runtime type in the top bar.
#!nvidia-smi

"""## Installation

All file changes made by this notebook are temporary. 
You can try to mount your own google drive to store files if you want.

"""

#(repositorio original)
#!git clone https://github.com/neuralchen/SimSwap 
#!cd SimSwap && git pull
#(mi repositorio con  GUI 512, opciones y preparado para 512)
#!git clone https://github.com/alu0100585704/simswap-main.git 
#!cd simswap-main && git pull

#!pip install insightface==0.2.1 onnxruntime moviepy
#!pip install googledrivedownloader
#!pip install imageio==2.4.1

#!pip list

import os
#os.chdir("SimSwap")
#os.chdir("simswap-main")

#!ls

#from google_drive_downloader import GoogleDriveDownloader

### it seems that google drive link may not be permenant, you can find this ID from our open url.
# GoogleDriveDownloader.download_file_from_google_drive(file_id='1TLNdIufzwesDbyr_nVTR7Zrx9oRHLM_N',
#                                     dest_path='./arcface_model/arcface_checkpoint.tar')
# GoogleDriveDownloader.download_file_from_google_drive(file_id='1PXkRiBUYbu1xWpQyDEJvGKeqqUFthJcI',
#                                     dest_path='./checkpoints.zip')

#!wget -P ./arcface_model https://github.com/neuralchen/SimSwap/releases/download/1.0/arcface_checkpoint.tar
#!wget https://github.com/neuralchen/SimSwap/releases/download/1.0/checkpoints.zip
#!unzip ./checkpoints.zip  -d ./checkpoints
#!wget https://github.com/neuralchen/SimSwap/releases/download/512_beta/512.zip
#!unzip ./512.zip  -d ./checkpoints
#!wget -P ./parsing_model/checkpoint https://github.com/neuralchen/SimSwap/releases/download/1.0/79999_iter.pth

## You can upload filed manually
# from google.colab import drive
# drive.mount('/content/gdrive')

### Now onedrive file can be downloaded in Colab directly!
### If the link blow is not permanent, you can just download it from the 
### open url(can be found at [our repo]/doc/guidance/preparation.md) and copy the assigned download link here.
### many thanks to woctezuma for this very useful help
#!wget --no-check-certificate "https://sh23tw.dm.files.1drv.com/y4mmGiIkNVigkSwOKDcV3nwMJulRGhbtHdkheehR5TArc52UjudUYNXAEvKCii2O5LAmzGCGK6IfleocxuDeoKxDZkNzDRSt4ZUlEt8GlSOpCXAFEkBwaZimtWGDRbpIGpb_pz9Nq5jATBQpezBS6G_UtspWTkgrXHHxhviV2nWy8APPx134zOZrUIbkSF6xnsqzs3uZ_SEX_m9Rey0ykpx9w" -O antelope.zip
#!unzip ./antelope.zip -d ./insightface_func/models/

"""## Inference"""

import cv2
import torch
import fractions
import numpy as np
from PIL import Image
import torch.nn.functional as F
from torchvision import transforms
from models.models import create_model
from options.test_options import TestOptions
from insightface_func.face_detect_crop_multi import Face_detect_crop
from util.videoswap import video_swap
from util.add_watermark import watermark_image

transformer = transforms.Compose([
        transforms.ToTensor(),
        #transforms.Normalize([0.485, 0.456, 0.406], [0.229, 0.224, 0.225])
    ])

transformer_Arcface = transforms.Compose([
        transforms.ToTensor(),
        transforms.Normalize([0.485, 0.456, 0.406], [0.229, 0.224, 0.225])
    ])

detransformer = transforms.Compose([
        transforms.Normalize([0, 0, 0], [1/0.229, 1/0.224, 1/0.225]),
        transforms.Normalize([-0.485, -0.456, -0.406], [1, 1, 1])
    ])

opt = TestOptions()
opt.initialize()
opt.parser.add_argument('-f') ## dummy arg to avoid bug
opt = opt.parse()

opt.pic_a_path = './demo_file/Iron_man.jpg' ## or replace it with image from your own google drive
opt.video_path = './demo_file/multi_people_1080p.mp4' ## or replace it with video from your own google drive
opt.output_path = './output/demo.mp4'
opt.temp_path = './tmp'
opt.Arc_path = './arcface_model/arcface_checkpoint.tar'
opt.isTrain = False
opt.use_mask = True  ## new feature up-to-date
opt.crop_size = 224

pic_a_path = input("Imagen origen por defecto: " + opt.pic_a_path + "\nIntroduzca ruta con nomnbre de nueva imagen o presiona ENTER para aceptar: ") 
if len(pic_a_path) > 0:    
    opt.pic_a_path = pic_a_path

video_path = input("Video origen por defecto: " + opt.video_path + "\nIntroduzca ruta con nombre de nuevo Video o presiona ENTER para aceptar: ") 
if len(video_path) > 0:    
    opt.video_path = video_path


output_path = input("Video destino por defecto: " + opt.output_path + "\nIntroduzca ruta con nombre de nuevo video de destiono o presiona ENTER para aceptar: ") 
if len(output_path) > 0:    
    opt.output_path = output_path

crop_size = input("Crop Size por defecto: " + str(opt.crop_size) + "\nIntroduzca nuevo valor(512 por ejemplo) o presiona ENTER para aceptar: ") 
if len(crop_size) > 0:    
    opt.crop_size = int(crop_size)


if len(input("Quitar marca de agua S/N")) > 0:    
    opt.no_simswaplogo = True


torch.nn.Module.dump_patches = True
model = create_model(opt)
model.eval()

app = Face_detect_crop(name='antelope', root='./insightface_func/models')
app.prepare(ctx_id= 0, det_thresh=0.6, det_size=(640,640))

with torch.no_grad():
    pic_a = opt.pic_a_path
    # img_a = Image.open(pic_a).convert('RGB')
    img_a_whole = cv2.imread(pic_a)
    img_a_align_crop, _ = app.get(img_a_whole,crop_size)
    img_a_align_crop_pil = Image.fromarray(cv2.cvtColor(img_a_align_crop[0],cv2.COLOR_BGR2RGB)) 
    img_a = transformer_Arcface(img_a_align_crop_pil)
    img_id = img_a.view(-1, img_a.shape[0], img_a.shape[1], img_a.shape[2])

    # convert numpy to tensor
    img_id = img_id.cuda()

    #create latent id
    img_id_downsample = F.interpolate(img_id, size=(112,112))
    latend_id = model.netArc(img_id_downsample)
    latend_id = latend_id.detach().to('cpu')
    latend_id = latend_id/np.linalg.norm(latend_id,axis=1,keepdims=True)
    latend_id = latend_id.to('cuda')

    video_swap(opt.video_path, latend_id, model, app, opt.output_path, temp_results_dir=opt.temp_path, use_mask=opt.use_mask)

