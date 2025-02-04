import tkinter as tk
from tkinter import messagebox
import threading
import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
from urllib.parse import urlparse
import datetime
import re


def collect_data():
    """
    NTIS, SMTECH, NIA, KOITA 네 사이트를 순차적으로 크롤링하여
    '사이트', '공고명', '마감일', '현황' 정보를 추출한 후,
    announcements.xlsx 파일로 저장하는 함수.

    NTIS URL: https://www.ntis.go.kr/rndgate/eg/un/ra/mng.do?pageIndex={page}  (1~10페이지)
    SMTECH URL: https://www.smtech.go.kr/front/ifg/no/notice02_list.do?pageIndex={page}  (1~10페이지)
    NIA URL: https://nia.or.kr/site/nia_kor/ex/bbs/List.do?cbIdx=78336&pageIndex={page}  (1~3페이지)
    KOITA URL: https://www.koita.or.kr/board/commBoardNotice003List.do?page={page}&  (목록에서 최대 20개, 보통 1~2페이지)
    """
    data = []

    # ---------------- NTIS 사이트 크롤링 ----------------
    ntis_base_url = 'https://www.ntis.go.kr/rndgate/eg/un/ra/mng.do'
    for page in range(1, 11):
        url = f'{ntis_base_url}?pageIndex={page}'
        print(f"NTIS 페이지 {page} 스크래핑 중... URL: {url}")
        domain = urlparse(url).netloc

        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            table = soup.find('table', class_='basic_list')
            if table:
                tbody = table.find('tbody')
                rows = tbody.find_all('tr')
                for row in rows:
                    status_td = row.find('td', {'data-title': '현황'})
                    title_td = row.find('td', {'data-title': '공고명'})
                    deadline_td = row.find('td', {'data-title': '마감일'})
                    status = status_td.get_text(strip=True) if status_td else ''
                    title = title_td.get_text(strip=True) if title_td else ''
                    deadline = deadline_td.get_text(strip=True) if deadline_td else ''
                    data.append({
                        '사이트': domain,
                        '공고명': title,
                        '마감일': deadline,
                        '현황': status
                    })
            else:
                print(f"NTIS 페이지 {page}에서 기본 리스트 테이블(클래스 'basic_list')을 찾을 수 없습니다.")
        else:
            print(f"NTIS 페이지 {page}에 접근할 수 없습니다. 상태 코드: {response.status_code}")
        time.sleep(0.1)

    # ---------------- SMTECH 사이트 크롤링 ----------------
    smtech_base_url = 'https://www.smtech.go.kr/front/ifg/no/notice02_list.do'
    for page in range(1, 11):
        url = f'{smtech_base_url}?pageIndex={page}'
        print(f"SMTECH 페이지 {page} 스크래핑 중... URL: {url}")
        domain = urlparse(url).netloc

        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            # SMTECH 목록 테이블의 클래스는 "tbl_base tbl_type01"
            table = soup.find('table', class_='tbl_base tbl_type01')
            if table:
                tbody = table.find('tbody')
                rows = tbody.find_all('tr')
                for row in rows:
                    tds = row.find_all('td')
                    if len(tds) < 6:
                        continue
                    title_text = tds[2].get_text(strip=True)
                    period_text = tds[3].get_text(strip=True)
                    if '~' in period_text:
                        deadline_text = period_text.split('~')[-1].strip()
                    else:
                        deadline_text = period_text
                    status_td = tds[5]
                    img = status_td.find('img')
                    if img and img.has_attr('alt'):
                        status_text = img['alt'].strip()
                    else:
                        status_text = status_td.get_text(strip=True)
                    data.append({
                        '사이트': domain,
                        '공고명': title_text,
                        '마감일': deadline_text,
                        '현황': status_text
                    })
            else:
                print(f"SMTECH 페이지 {page}에서 테이블(tbl_base tbl_type01)을 찾을 수 없습니다.")
        else:
            print(f"SMTECH 페이지 {page}에 접근할 수 없습니다. 상태 코드: {response.status_code}")
        time.sleep(0.1)

    # ---------------- NIA 사이트 크롤링 ----------------
    nia_base_url = 'https://nia.or.kr/site/nia_kor/ex/bbs/List.do?cbIdx=78336'
    for page in range(1, 4):
        url = f'{nia_base_url}&pageIndex={page}'
        print(f"NIA 페이지 {page} 스크래핑 중... URL: {url}")
        domain = urlparse(url).netloc

        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            board_div = soup.find("div", class_="board_type01")
            if board_div:
                items = board_div.find_all("li")
                for item in items:
                    subject_tag = item.find("span", class_="subject searchItem")
                    subject_text = subject_tag.get_text(strip=True) if subject_tag else ""
                    src_tag = item.find("span", class_="src")
                    if src_tag:
                        deadline_text = src_tag.get_text(strip=True)
                        if not deadline_text.endswith("~"):
                            deadline_text += "~"
                    else:
                        deadline_text = ""
                    status_text = "확인바람"
                    data.append({
                        '사이트': domain,
                        '공고명': subject_text,
                        '마감일': deadline_text,
                        '현황': status_text
                    })
            else:
                print(f"NIA 페이지 {page}에서 board_type01 컨테이너를 찾을 수 없습니다.")
        else:
            print(f"NIA 페이지 {page}에 접근할 수 없습니다. 상태 코드: {response.status_code}")
        time.sleep(0.1)

    # ---------------- KOITA 사이트 크롤링 ----------------
    # KOITA 목록은 페이지 1과 2에서 최대 20개(예: 순번 563부터 544까지)의 데이터를 수집합니다.
    koita_items_count = 0
    max_koita_items = 20
    for page in [1, 2]:
        url = f'https://www.koita.or.kr/board/commBoardNotice003List.do?page={page}&'
        print(f"KOITA 페이지 {page} 스크래핑 중... URL: {url}")
        domain = urlparse(url).netloc

        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            table = soup.find("table", class_="tb tb_col tb_bd tb_st01")
            if table:
                tbody = table.find("tbody")
                rows = tbody.find_all("tr")
                for row in rows:
                    if koita_items_count >= max_koita_items:
                        break
                    a_tag = row.find("a", href=lambda href: href and "page_move" in href)
                    if a_tag:
                        href_value = a_tag.get("href")
                        match = re.search(r"{\s*no:\s*(\d+)\s*}", href_value)
                        if match:
                            no_value = match.group(1)
                            detail_url = f"https://www.koita.or.kr/board/commBoardNotice003View.do?page={page}&no={no_value}"
                        else:
                            continue
                        print(f"KOITA 상세페이지 스크래핑: {detail_url}")
                        detail_response = requests.get(detail_url)
                        if detail_response.status_code == 200:
                            detail_soup = BeautifulSoup(detail_response.text, 'html.parser')
                            # 공고명: 요소(class="f_st01") 내 텍스트
                            title_tag = detail_soup.find(class_="f_st01")
                            title_text = title_tag.get_text(strip=True) if title_tag else ""
                            # 마감일: "공고기간" 아래 <td>의 <div>의 텍스트에서 "~" 이후의 날짜 추출
                            deadline_text = ""
                            th_tags = detail_soup.find_all("th")
                            for th in th_tags:
                                if "공고기간" in th.get_text():
                                    td = th.find_next_sibling("td")
                                    if td:
                                        div_tag = td.find("div")
                                        if div_tag:
                                            period_str = div_tag.get_text(strip=True)
                                            if "~" in period_str:
                                                deadline_text = period_str.split("~")[-1].strip()
                                            else:
                                                deadline_text = period_str.strip()
                                    break
                            # 만약 마감일이 빈칸이면 그대로 빈칸 처리
                            # 현황 결정: 마감일이 있으면 날짜형태(YYYY-MM-DD)로 변환하여 오늘과 비교
                            if deadline_text:
                                try:
                                    deadline_date = datetime.datetime.strptime(deadline_text, "%Y-%m-%d").date()
                                    today = datetime.date.today()
                                    if deadline_date >= today:
                                        status_text = "접수중"
                                    else:
                                        status_text = "접수완료"
                                except Exception as e:
                                    status_text = "확인바람"
                            else:
                                status_text = "확인바람"
                            data.append({
                                '사이트': urlparse(detail_url).netloc,
                                '공고명': title_text,
                                '마감일': deadline_text,
                                '현황': status_text
                            })
                            koita_items_count += 1
                        else:
                            print(f"KOITA 상세페이지에 접근할 수 없습니다. 상태 코드: {detail_response.status_code}")
                        time.sleep(0.1)
                    else:
                        continue
                if koita_items_count >= max_koita_items:
                    break
            else:
                print(f"KOITA 페이지 {page}에서 테이블(tb tb_col tb_bd tb_st01)을 찾을 수 없습니다.")
        else:
            print(f"KOITA 페이지 {page}에 접근할 수 없습니다. 상태 코드: {response.status_code}")
        time.sleep(0.1)
        if koita_items_count >= max_koita_items:
            break

    # ---------------- Excel 파일로 저장 ----------------
    df = pd.DataFrame(data)
    excel_filename = 'announcements.xlsx'
    df.to_excel(excel_filename, index=False)
    messagebox.showinfo("완료", f"엑셀 파일로 저장되었습니다: {excel_filename}")


def on_collect():
    """'정보수집' 버튼 클릭 시 크롤링 작업을 별도 스레드로 실행"""
    threading.Thread(target=collect_data).start()


def on_exit():
    """'끝내기' 버튼 클릭 시 프로그램 종료"""
    root.destroy()
    root.quit()
    exit()


# Tkinter UI 창 생성 (UI는 기존대로 유지)
root = tk.Tk()
root.title("정보 수집 프로그램")
root.geometry("300x150")

collect_button = tk.Button(root, text="정보수집", command=on_collect, width=15, height=2)
collect_button.pack(pady=20)

exit_button = tk.Button(root, text="끝내기", command=on_exit, width=15, height=2)
exit_button.pack(pady=10)

root.mainloop()
