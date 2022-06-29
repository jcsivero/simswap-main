import torch
device = torch.device("cuda:0" if torch.cuda.is_available() else "cpu")
if torch.cuda.device_count() > 0:
    print (torch.cuda.current_device())
    print(torch.cuda.get_device_name(0))

print(device)

