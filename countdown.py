from progress.bar import IncrementalBar
from progress.counter import Countdown
import random
import time
def sleep():
    t = 0.01
    t += t * random.uniform(0.001, 30)  # Add some variance
    time.sleep(t)

# bar =  Countdown('Программа закроется через: ', suffix=' сек.')
for i in Countdown('Закроюсь через: ').iter(range(60)):
    sleep()