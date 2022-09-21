import time
from operator import itemgetter
import vk_api
import datetime
from time import strptime

import xlsxwriter as xlsxwriter


# _______________________________________________________________________________________________________________


def captcha_handler(captcha):
    key = input("Enter captcha code {0}: ".format(captcha.get_url())).strip()
    return captcha.try_again(key)


def get_group_comments(vk, post, user):
    post_comments = list()
    offset = 0
    msg_comments = post['comments']
    msg_num = msg_comments['count']

    if msg_num == 0:
        return post_comments.__dict__

    while msg_num > 0:
        try:
            post_comments_api = vk.wall.getComments(owner_id=post['from_id'], post_id=post['id'], offset=offset,
                                                    count=100)
        except vk_api.exceptions.ApiError as error_msg:
            print(error_msg)
            msg_num = -1
            break

        offset += 100
        msg_num -= offset
        for item in post_comments_api['items']:

            if user['id'] != item['from_id']:
                continue

            post_comments.append({
                'reply_text': item['text'],
                'group_id': -post['owner_id'],
                'post_id': item['post_id'],
                'post_date': post['date'],
                'post_text': post['text'],
                'reply_id': item['id'],
                'reference': f"https://vk.com/wall{post['owner_id']}_{item['post_id']}?reply={item['id']}",
                'user_id': item['from_id'],
                'reply_date': item['date']
            })
            print('Comment: ' + item['text'])
    return post_comments


def main():
    posts = list()
    result_list = list()
    from_date = 1663343909

    vk_session = vk_api.VkApi('mail@mail.xxx', 'password', captcha_handler=captcha_handler)

    try:
        vk_session.auth()
    except vk_api.exceptions.Captcha as captcha:
        sid = captcha.sid  # Get sid
        print(sid)
        captcha.get_url()  # Get ref of captcha image
        captcha.get_image()  # Get captcha image (jpg)

    try:
        vk_session.auth(token_only=True)
    except vk_api.AuthError as error_msg:
        print(error_msg)
        return

    vk = vk_session.get_api()

    user_id = input('Enter UID:')
    try:
        user = vk.users.get(user_ids=user_id, fields='city,connections,contacts,country,domain,education,screen_name')
        print(user)
    except vk_api.AccessDenied as error_msg:
        print(error_msg)
        return

    date = datetime.datetime.fromtimestamp(from_date)
    ans = input(f"\nDefault date is {date.date().strftime('%d.%m.%Y').__str__()} ({from_date})"
                f" \nDo you want to change it?")
    if ans == 'Y' or ans == 'y':
        time_new = strptime(input('Enter new date:'), '%d.%m.%Y')
        from_date = int(time.mktime(time_new))
        print(f"New timestamp is {from_date}")

    group_set = set()
    with open(r'C:\Users\ibaro\Downloads\target.txt', 'r', encoding='utf-8', errors='ignore') as fl:
        file_list = fl.readlines()
        group_list = vk.groups.getById(group_ids=','.join(file_list))
        for group in group_list:
            group_set.add(group['id'])

    # user_groups = set()
    try:
        user_groups = set(vk.groups.get(user_id=user[0]['id'])['items'])
    except vk_api.exceptions.ApiError as error_msg:
        print(error_msg)
        user_groups = set()

    group_set.update(user_groups)

    if group_set.__len__ == 0:
        print("No groups were specified!!!")

    for group in group_set:
        print(f"Checking group {group}")
        offset = 0
        finish_search = ''
        while True:
            try:
                group_posts = sorted(vk.wall.get(owner_id=-group, offset=offset, count=100)['items'],
                                     key=itemgetter('date'), reverse=True)
            except vk_api.exceptions.ApiError as error_msg:
                print(error_msg)
                break

            post_num = offset

            for post in group_posts:
                post_num += 1
                date = datetime.datetime.fromtimestamp(post['date'])
                print(f"Post({post_num}) {post['id']} datetime: {date.date().strftime('%d.%m.%Y').__str__()} "
                      f"({post['date']}) \n")

                post_text = post['text'].replace("\n\n", "\n")
                post_text.replace("\n\n\n", "\n")
                post_text.replace("\n\n", "\n")

                print(f"{post_text} \n\n")
                if post['date'] < from_date:
                    finish_search = 'X'
                    break
                else:
                    msg_comments = post['comments']
                    msg_num = msg_comments['count']

                    if msg_num != 0:
                        posts.append(post)

                        results = get_group_comments(vk, post, user[0])

                        for result in results:
                            result_list.append(result)

            if finish_search == 'X':
                break
            offset += 100
    for result in result_list:
        print(result)

    export_to_xls(result_list)


def export_to_xls(data_list):
    if data_list.__len__() == 0:
        print('No data to convert!')
        return

    workbook = xlsxwriter.Workbook('sample_data1.xlsx')  # here it's better to define your rules of naming files
    worksheet = workbook.add_worksheet()

    row = 0

    for line in data_list:

        col = 0

        for field in line:
            worksheet.write(row, col, line.get(field))
            col += 1
        row += 1

    workbook.close()


if __name__ == '__main__':
    main()
