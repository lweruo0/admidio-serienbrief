import os
import hashlib
import json
import urllib.request
import datetime


def get_otp_liste():
    """
    hole neue otp hashpasswords
    """
    password = "acAmQNucVlWMk7TlmHwWnlk8IXcLpnYv8lq17"
    datehash = hashlib.md5(datetime.datetime.now().strftime("%Y%m%d").encode("ascii")).hexdigest()
    wp_url = f"https://bfv-ehingen.de/?getOTPlist=1&passwd={password}&nonce={datehash}&clearotp=1"

    print(wp_url)
    with urllib.request.urlopen(wp_url) as req:
        contents = req.read()

    import pprint

    try:
        inhalt = json.loads(contents)
    except:
        inhalt = "Anfrage gesperrt"

    json.dump(inhalt, open("otp.json", "w", encoding="utf-8"))

    pprint.pprint(inhalt["BenKunter"])

    # print https://bfv-ehingen.de/mitgliedsdaten?mgn=6&otp=
    for v in inhalt.values():
        mgn = v[1]
        login = v[0]
        otp = v[2]
        url = f"https://bfv-ehingen.de/gewinnspiel?nr={mgn}&ot={otp}"
        if login == "BenKunter":
            print(url)

    # https://bfv-ehingen.de/mitgliedsdaten?mgn=6&otp=xXBRTpdUcJ8i0hUcn8theCZ1EeG0n7


if __name__ == "__main__":
    get_otp_liste()
# https://bfv-ehingen.de/?getOTPlist=1&passwd=acAmQNucVlWMk7TlmHwWnlk8IXcLpnYv8lq17&nonce=c742cbe5f2893cbe578ffa442c07cc17
# ['BenKunter', '6', 'Ohocb5B65nzbl', 'zGI0MUSNvqfVe']
# https://bfv-ehingen.de/gewinnspiel?nr=6&ot=Ohocb5B65nzbl
