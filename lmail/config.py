import time
log_path = time.strftime('./mail_%y_%m_%d.log')
log_name = 'mail'
default_sender = 'example@hust.edu.cn'
default_password = 'passwd'
default_host = 'mail.hust.edu.cn'
default_port = 25
mails_per_send = 50
mails_per_login = 300
interval_between_login = 1
reset_per_n_login = 4
reset_interval = 5
