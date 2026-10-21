from __future__ import annotations

import base64
import zlib


_CFB_SIGNATURE = bytes.fromhex("D0CF11E0A1B11AE1")
_PACKAGED_OLE_CFB_ZLIB_BASE64 = (
    "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"
    "zHZF/KLy3OuT6u01/PrGfd/w3GZC8anRFLNX6xuzGg+suXR1lym16c9TpdP7+wJuTRtTRqcusu5y6tvj8TTxgpFCZf3O4tIJJz6M"
    "sTJ9begfVvLIbXJXxYzMdQZ6nvNK1Hd+rLpYY+G2MW185u7gktU9S6IDTnll5SQa5Ry69c2deN3lV/O0H2lqLclT7uu9zsx1f9rk"
    "2/mcdPbkkFlefz//utP0Sc6ZKm/XSeUTfCkbTl/Ku/2m7OvYJuoY22vnh32yPWSVub9ldt20r/onX85luLS5lD2kf1DeqlV/UiUz"
    "/GY0Xd+mp9ZjT+abe3dmL6p5kHUqa43r3OZJl5k6EZ+8X2q3N6eGpBLRA03QXyV3Sl9cQl9ZE1grGsITGYuW2dmifXr2zbh6e5WN"
    "7zIubGhLmXKrXs/sV7uSHF+7gymg2yI1xtjwnuHre5Mb5ZWqKemH3vUd8fp5fMvX7vwXdOruo/L1T3y9r8d0/+pVm7q+KeNQqRkl"
    "ZMWSfxxcfa8yvCao7aBni4Nhybiz9ju01Dd1Ta4JPj5mQVuTxq0Jfvec55umluYuDe802rlH/7FrQPsnanCnW8Tx7LIm5ul6++Y9"
    "ky+nnrB+7kde7JYfEPts/v4yu6dTL+SUMY7eprVWnvt65wcBr6ILVxZQN8CjzewvRpJiGPEMO5594LzF86IZiShCVJoDlerpGqO7"
    "U/Yvz+lyi6y28Jt/6vrx6s/At9jAp3Hfrb33Op8/3eB245FOR9FE1dnWxQqqmxkttpt+f7/GutSw7O64wQ3Wi7obB/W8WqHXdGab"
    "9qPxV5dNSTHpyMmoqHxAv1Np1TH0nfWDuedN47y2+M982d5n+cTvtZW5EtEd3/i41z4HfWCP5DuddTcIGe+AzH4ze2KCoZvGRd2X"
    "qfTPD2flxXc5GiusvbpUuzbC0jLxkOmEiUXVoTn58Q82k+Xm7LMIotussmwNUfNrA6qt+wp/6pEU755h7vdNrDQ2G76WPrq97rPf"
    "xBSn7+5OxfotDy8uK7BPGpN+yLPC/3lICSX3TH2189vOgwoVFdkFJqaDJlXv7M2/fXBobttvHere57bNnBagtX70qB1W52gNExhW"
    "y5adIl2f9fCJ/vjuapqu4oz3Ft+Wha2b7JNQ1zO8JDjZ9rLztqDN2kHVj163qfmkpn5+XhS91/+YyVtV+ZBDE/7KUNMOrChaYSx/"
    "z8t2Yfrga6rmrYpRnyiuf31op46Yqjl/1vAvB2g7msOdf1+1Wu3EszEBLTnJz842h9sWKidlW0VNbkvQLSow2PE8I3H/I7Bz9xWX"
    "gk+p+2+CJaWkYZWhob89b4iLPZ8XOSTnh15rWMjwm/0r/JnfdV4kVlTNnf3+o2FOR619hl18h0HRFtrDNOu1S5VqZxwt7+zxqiHM"
    "XpOsVNs6K+Sos2cNoWHr4cb2l3/d+pDo4q+bQ4u+57OiYo6NTXzPewelA/s32V7wyy6Nm1Fu9mXTFhsQQ6/ffHicqtOt+iGUp5UB"
    "SzNdjzTr0cZa1FsWngw9mn9dY5VyyYTeBOusmZRfD7nYRRiXFjgbmXqsjzpKe0CqyCWt/PTh+denui+yU71SPzz73nszZuPnSMb0"
    "lYn33N9WJybsSfnR2bf3gHfq1+6W9oS6Fa+ehDM/Nrx+tDa1aNW5uMipNd8+f32tezM8pa/DyHuUaky49rm3dSHOQ7Xe5k5Nuqrt"
    "vs8h4+beMw0dpnMcDu8hZ/ksH53s20BbkrBpNuX3SWu0R+nkMMZcqt5gkr3QiUryI3k8ytZWP3d9RMgL45i0/lV2jXUVPnfVE2zl"
    "bf/WeLq99aGh+6fB9/UW5MyNHfd8bIDnbgWVUqW+8RN0bj3qjRxDLF39y5HozrC7mqEf9ozJtpjT6ByTR41fvc6Zevew4heGnjHl"
    "+OdFp/s1KpZRR34yH9Gbu8TGXjknQbNnV68tzaI4fdrZJXWhG95FDNo4znXjgadPDILXlfanL739UsvS17GfeOO1+49R//iDmtRv"
    "a2g+Kthi1eBb5zeMWJ5b9H2FyZXx1DrXWIPqP7OfO6RbjFUIzXhcfb926Df972ntu/5eOy7zxd9btC+Txi2xLVma9jJB3vjx9er0"
    "1b33Bp+ime+3Pm2uZVd9XNFga9Ynn0fKb4eM0RvVEDamdUM2SeXOyp6uNQHf5qrMZhSlE1/TPkdO9y79RWvBt92vdM77luieCrz1"
    "wunLG7VhWjc1ShWpp53fresMK3K7S7SNfFzcX77sDbHP4mOqhp3PwTcTW7vJfQEVCg6f4wrXFnSvGKJ0M7XwfOQSu2xPg8JzhXTf"
    "49akzAc7Vh6Ln7NzpeFW0uEMy7uKJSEavfe9qW3JHZtU1yd6MW9lPQDbcjvfmTpfKPvaVT/mfXtqaZ/iH0+rX62q23yhJrlhbfSM"
    "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"
    "85pmOfALZIjEbGBgV0A6e9z/C/UfxsP/HyaQVxk="
)


def default_ole_cfb_bytes() -> bytes:
    """Return a small, valid Compound File payload for OLE fallback objects."""

    return zlib.decompress(base64.b64decode(_PACKAGED_OLE_CFB_ZLIB_BASE64))


def normalize_hwpx_ole_payload(data: bytes) -> bytes:
    """Add the four-byte HWPX OLE length prefix to a raw CFB payload."""

    payload = bytes(data)
    if payload.startswith(_CFB_SIGNATURE):
        return len(payload).to_bytes(4, "little", signed=False) + payload
    return payload
