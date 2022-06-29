import torch
#pip install torch==1.10.1+cu113 torchvision==0.11.2+cu113 torchaudio===0.10.1+cu113 -f https://download.pytorch.org/whl/cu113/torch_stable.html
#install torch==1.11.0+cu113 -f https://download.pytorch.org/whl/cu113/torch_stable.html
#install torchvision==0.13.0+cu113 -f https://download.pytorch.org/whl/cu113/torch_stable.html
#install torchaudio===0.12.0+cu113 -f https://download.pytorch.org/whl/cu113/torch_stable.html

#install torch==1.12.0+cu116 -f https://download.pytorch.org/whl/cu116/torch_stable.html
#install torchvision==0.13.0+cu116 -f https://download.pytorch.org/whl/cu116/torch_stable.html
#install torchaudio===0.12.0+cu116 -f https://download.pytorch.org/whl/cu116/torch_stable.html

print (torch.cuda.is_available())
device = torch.device("cuda:0" if torch.cuda.is_available() else "cpu")
if torch.cuda.device_count() > 0:
    print (torch.cuda.current_device())
    print(torch.cuda.get_device_name(0))

print(device)

