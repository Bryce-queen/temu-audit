import sys
import os

# 指向你项目的路径
path = '/home/LINGH/temu-audit'
if path not in sys.path:
    sys.path.append(path)

# 在这里填入你的 Stripe 测试密钥（去 https://dashboard.stripe.com/test/apikeys 拿）
os.environ['STRIPE_SECRET_KEY'] = 'sk_test_xxxxxxxxxxxxxxxxxx'
os.environ['STRIPE_PUBLISHABLE_KEY'] = 'pk_test_xxxxxxxxxxxxxxxxxx'
os.environ['PUBLIC_DOMAIN'] = 'https://LINGH.pythonanywhere.com'

from app import app as application
