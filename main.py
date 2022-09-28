from utils import *


# _______________________________________________________________________________________________________________


def get_group_comments(vk, post):
    post_comments_api = list()
    threads_comments_api = list()

    offset = 0
    msg_comments = post['comments']
    comment_count = post['comments']['count']

    if comment_count == 0:
        return post_comments_api.__dict__

    while comment_count > offset:
        try:
            post_comments_api = vk.wall.getComments(owner_id=post['from_id'],
                                                    post_id=post['id'],
                                                    offset=offset,
                                                    count=100)['items']

            if comment_count > 100:
                offset += 100
            else:
                offset += post_comments_api.__len__()

        except vk_api.exceptions.ApiError as error_msg:
            print(error_msg)
            break

        for item in post_comments_api:
            offset_thread = 0
            thread_count = item['thread']['count']
            if thread_count != 0:
                while thread_count > offset_thread:
                    thread_comments_api = vk.wall.getComments(owner_id=post['from_id'],
                                                              post_id=post['id'],
                                                              offset=offset_thread,
                                                              count=100,
                                                              comment_id=item['id'])['items']
                    if thread_count > 100:
                        offset_thread += 100
                    else:
                        offset_thread += thread_comments_api.__len__()

                    for comment in thread_comments_api:
                        threads_comments_api.append(comment)
                        offset += 1

        for thread in threads_comments_api:
            post_comments_api.append(thread)

    return post_comments_api


def main():
    posts = list()
    user_comments = list()

    vk = authorize()
    from_date = int(time.time())
    user = define_user(vk)[0]
    from_date = define_date(from_date)
    group_set = define_groups(vk, user)

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

                        comments = get_group_comments(vk, post)
                        user_comments_post = filter_comments(comments, user, post)

                        for comment in user_comments_post:
                            user_comments.append(comment)

            if finish_search == 'X':
                break
            offset += 100
    for comment in user_comments:
        print(comment)

    export_to_xls(user_comments, user['domain'])


if __name__ == '__main__':
    main()
