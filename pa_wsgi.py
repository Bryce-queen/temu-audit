import sys
import os

# 指向你项目的路径
path = '/home/LINGH/temu-audit'
if path not in sys.path:
    sys.path.append(path)

# 在这里填入你的 Stripe 测试密钥（去 https://dashboard.stripe.com/test/apikeys 拿）
os.environ['STRIPE_SECRET_KEY'] = 'sk_test_xxxxxxxxxxxxxxxxxx'
os.environ['STRIPE_PUBLISHABLE_KEY'] = 'pk_test_xxxxxxxxxxxxxxxxxx'
os.environ['STRIPE_WEBHOOK_SECRET'] = 'whsec_xxxxxxxxxxxxxxxxxx'
os.environ['ADMIN_TOKEN'] = 'change_me_to_a_long_random_value'
os.environ['SECRET_KEY'] = 'change_me_to_another_long_random_value'
os.environ['ZHIPU_API_KEY'] = ''
os.environ['SILICONFLOW_API_KEY'] = ''
os.environ['COZE_BOT_ID'] = ''
os.environ['COZE_TOKEN'] = ''
os.environ['PUBLIC_DOMAIN'] = 'https://LINGH.pythonanywhere.com'

from app import app as application
