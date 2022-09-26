import time
from operator import itemgetter

import vk_api
import datetime
from time import strptime

import xlsxwriter as xlsxwriter


# _______________________________________________________________________________________________________________


def get_group_comments(vk, post, user):
    post_comments = list()
    threads_comments_api = list()
    offset = 0
    msg_comments = post['comments']
    comment_count = msg_comments['count']

    if comment_count == 0:
        return post_comments.__dict__

    while comment_count > offset:
        try:
            post_comments_api = vk.wall.getComments(owner_id=post['from_id'],
                                                    post_id=post['id'],
                                                    offset=offset,
                                                    count=100)
        except vk_api.exceptions.ApiError as error_msg:
            print(error_msg)
            break

        for item in post_comments_api['items']:
            offset_thread = 0
            if item['thread']['count'] != 0:
                while item['thread']['count'] > offset_thread:
                    thread_comments_api = vk.wall.getComments(owner_id=post['from_id'],
                                                              post_id=post['id'],
                                                              offset=offset_thread,
                                                              count=100,
                                                              comment_id=item['id'])['items']
                    for comment in thread_comments_api:
                        threads_comments_api.append(comment)

                    offset_thread += 100

        for thread in threads_comments_api:
            post_comments_api['items'].append(thread)

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
        offset += 100
    return post_comments


def main():
    posts = list()
    result_list = list()
    user_groups = set()

    from_date = int(time.time())

    vk = authorize()

    user_id = input('Enter UID:\n')
    try:
        user = vk.users.get(user_ids=user_id, fields='city,connections,contacts,country,domain,education,screen_name')
        print(user)
    except vk_api.AccessDenied as error_msg:
        print(error_msg)
        return

    date = datetime.datetime.fromtimestamp(from_date)

    try:
        ans = strptime(input(f"\nDefault from-date is {date.date().strftime('%d.%m.%Y').__str__()} ({from_date})"
                             f" \nInput new date to change it, otherwise input 'n'\n"), '%d.%m.%Y')

        from_date = int(time.mktime(ans))
        print(f"New timestamp is {from_date}")
    except ValueError:
        print("Date was not changed")

    group_set = set()
    try:
        with open(r'groups.txt', 'r', encoding='utf-8', errors='ignore') as groups_file:
            file_list = groups_file.readlines()
            if len(file_list):
                group_list = vk.groups.getById(group_ids=','.join(file_list))
                for group in group_list:
                    group_set.add(group['id'])
    except FileNotFoundError:
        open('groups.txt', mode='w')

    ans = input("\nDo you want to add user profile groups?\n")
    if ans == 'Y' or ans == 'y':
        try:
            user_groups = set(vk.groups.get(user_id=user[0]['id'])['items'])
        except vk_api.exceptions.ApiError as error_msg:
            print(error_msg)
            user_groups = set()

    group_set.update(user_groups)

    if not len(group_set):
        print("\nNo groups were specified!!! Add groups to catalog 'groups.txt'!!!")
        return

    for group in group_set:
        print(f"Checking group {group}")
        offset = 0
        finish_search = ''
        while True:
            try:
                group_posts = vk.wall.get(owner_id=-group, offset=offset, count=100)['items']
            except vk_api.exceptions.ApiError as error_msg:
                print(error_msg)
                break

            post_num = offset

            for post in group_posts:
                post_num += 1
                post_date = datetime.datetime.fromtimestamp(post['date'])

                if post_num == 1 and post['date'] < from_date:
                    continue

                print(f"Post({post_num}) {post['id']} datetime: {post_date.date().strftime('%d.%m.%Y').__str__()} "
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
                    comment_count = msg_comments['count']

                    if comment_count != 0:
                        posts.append(post)

                        results = get_group_comments(vk, post, user[0])

                        for result in results:
                            result_list.append(result)

            if finish_search == 'X':
                break
            offset += 100
    for result in result_list:
        print(result)

    export_to_xls(result_list, user_id)


def export_to_xls(data_list, user_id):
    if data_list.__len__() == 0:
        print('No data to convert!')
        return

    now = datetime.datetime.now()
    dt_string = now.strftime("%Y%m%d-%H%M%S")

    workbook = xlsxwriter.Workbook(f"{user_id}_{dt_string}.xlsx")
    worksheet = workbook.add_worksheet()

    row = 0

    for line in data_list:

        col = 0

        for field in line:
            worksheet.write(row, col, line.get(field))
            col += 1
        row += 1

    workbook.close()


def authorize():

    try:
        with open(r'login.txt', 'r', encoding='utf-8', errors='ignore') as login_file:
            login = login_file.readlines()[0]
    except FileNotFoundError:
        login = input("\nEnter login: ")

    try:
        with open(r'password.txt', 'r', encoding='utf-8', errors='ignore') as password_file:
            password = password_file.readlines()[0]
    except FileNotFoundError:
        password = input("\nEnter password: ")

    vk_session = vk_api.VkApi(login, password, captcha_handler=captcha_handler)

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

    open('login.txt', mode='w').write(login)
    open('password.txt', mode='w').write(password)

    return vk_session.get_api()


def captcha_handler(captcha):
    key = input("Enter captcha code {0}: ".format(captcha.get_url())).strip()
    return captcha.try_again(key)


if __name__ == '__main__':
    main()
