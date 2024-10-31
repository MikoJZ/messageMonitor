import schedule
import datetime
import time
from utils import play_random_audio


def remind_claim_monday():
    print("It's time to fill out your claim for Monday!")
    play_random_audio("../audio/monday_claim")

def remind_claim_friday():
    print("It's time to fill out your claim for Friday!")
    play_random_audio("../audio/friday_claim")

def remind_claim_monthly():
    if is_last_day_of_month():
        print("It's time to fill out your claim for the last day of the month!")
        play_random_audio("../audio/end_of_month_claim")

def is_last_day_of_month():
    today = datetime.datetime.now()
    next_day = today + datetime.timedelta(days=1)
    return today.month != next_day.month

def schedule_reminders():
    # 每周一的17:00提醒
    schedule.every().monday.at("17:00").do(remind_claim_monday)
    # 每周五的17:00提醒
    schedule.every().friday.at("17:00").do(remind_claim_friday)
    #schedule.every().thursday.at("10:24").do(remind_claim_friday)

    # 每个月的最后一天的17:00提醒
    schedule.every().day.at("17:00").do(remind_claim_monthly)

def run_reminders():
    schedule.run_pending()
