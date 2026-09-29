import { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import {
  ComposedChart, Bar, Line, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer, PieChart, Pie, Cell, BarChart, LabelList,
} from "recharts";
import {
  TrendingUp, Users, Target, MessageCircle, Pencil, Save, RotateCcw,
  X, ShoppingBag, CreditCard, CalendarDays, Sparkles, ChevronRight, Info,
  Briefcase, Plus, Trash2, Factory, Repeat,
} from "lucide-react";

/* ---------------------------------- design tokens ---------------------------------- */
const COLORS = {
  bg: "#FAF6EE",
  surface: "#FFFFFF",
  border: "#E7DFCF",
  ink: "#2A2420",
  inkSoft: "#6E6252",
  inkFaint: "#9A8F7B",
  teal: "#2F6E68",
  tealSoft: "#DCEAE7",
  ochre: "#D69A2D",
  ochreSoft: "#F6E7C6",
  clay: "#C1553B",
  claySoft: "#F3DAD2",
  slate: "#516B85",
  olive: "#8A8534",
  plum: "#8C5383",
};

const CHANNEL_META = {
  pos: { label: "SDC Gift Shop (POS)", color: COLORS.teal },
  custom: { label: "Custom Orders", color: COLORS.ochre },
  apoyo: { label: "Apoyo Sales", color: COLORS.clay },
  bold: { label: "BOLD in Africa", color: COLORS.slate },
  artisan: { label: "Artisan Katale", color: COLORS.olive },
  yujo: { label: "Yujo", color: COLORS.plum },
};
const CHANNEL_KEYS = Object.keys(CHANNEL_META);
// Apoyo products are consignment sales already included in POS revenue, so summing separately would double-count them
const REVENUE_KEYS = CHANNEL_KEYS.filter((k) => k !== "apoyo");
const GIFTSHOP_KEYS = REVENUE_KEYS.filter((k) => k !== "custom");

const MONTH_LABELS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const QUARTER_LABELS = ["Q1 (Jan-Mar)", "Q2 (Apr-Jun)", "Q3 (Jul-Sep)", "Q4 (Oct-Dec)"];
const CUSTOMER_CATEGORIES = ["Hotel/Restaurant", "Government Office", "NGO/Institution", "Individual", "Other"];
const CUSTOMER_CATEGORY_COLORS = { "Hotel/Restaurant": COLORS.teal, "Government Office": COLORS.slate, "NGO/Institution": COLORS.olive, "Individual": COLORS.ochre, "Other": COLORS.plum, "Unspecified": COLORS.inkFaint };

// 기본 시드 데이터 — 2026년(진행중) + 2025년(전년, priorYears) 실적을 실제 업로드 파일 기준으로 반영
// 새 엑셀을 업로드하면 이 값들은 갱신되고, 매년 파일의 감지된 연도에 따라 currentYear/priorYears로 자동 분류됨
const DEFAULT_DATA = {
  "channels": {
    "pos": [
      1521499.99,
      2394799.98,
      6576900.48,
      3395800.024,
      4904300,
      4295100,
      2545600,
      7071150,
      2113100,
      0,
      0,
      0
    ],
    "custom": [
      0,
      1200000,
      1350000,
      2570500,
      5515500,
      6560460,
      2573250,
      2348500,
      6384750,
      0,
      0,
      0
    ],
    "apoyo": [
      212000,
      746000,
      1864700,
      1192100,
      1207800,
      831600,
      343000,
      1163300,
      0,
      0,
      0,
      0
    ],
    "bold": [
      113750,
      143000,
      487500,
      120250,
      0,
      172500,
      150000,
      187500,
      0,
      0,
      0,
      0
    ],
    "artisan": [
      170339,
      140678,
      525424,
      300847,
      181356,
      40678,
      44915,
      198305,
      0,
      0,
      0,
      0
    ],
    "yujo": [
      149100,
      173600,
      499100,
      83300,
      37800,
      66500,
      0,
      87000,
      0,
      0,
      0,
      0
    ],
    "sawe": [
      0,
      0,
      0,
      0,
      0,
      9000,
      0,
      0,
      0,
      0,
      0,
      0
    ]
  },
  "visits": [
    50,
    74,
    236,
    85,
    73,
    76,
    81,
    108,
    35,
    0,
    0,
    0
  ],
  "contacts": [
    26,
    26,
    61,
    25,
    24,
    22,
    24,
    23,
    11,
    0,
    0,
    0
  ],
  "sold": [
    77,
    178,
    408,
    297,
    252,
    134,
    58,
    161,
    82,
    0,
    0,
    0
  ],
  "monthlyGoalUSD": 1500,
  "exchangeRate": 3500,
  "odmAnnualGoalUGX": 51000000,
  "consignmentVendors": [
    "Apoyo"
  ],
  "productGroups": [
    {
      "id": 1,
      "canonicalName": "T-shirt",
      "members": [
        "T-shirt",
        "T-shirt with one print extra",
        "T-shirt with more prints"
      ]
    }
  ],
  "itemVendorMap": {
    "4africazebra": "4Africa",
    "casualbag": "SDC",
    "bottlesticker": "SDC",
    "ribbonhairband(s)": "SDC",
    "puffhairband(s)": "SDC",
    "flowerkeyring": "SDC",
    "apoyohairband(s)": "Apoyo",
    "driedfruits(s)": "SDC",
    "tyde's1": "SDC",
    "turtle(l)": "SDC",
    "whale(s)": "SDC",
    "bookbagv2": "SDC",
    "apoyohairband(l)": "Apoyo",
    "picnicbag": "SDC",
    "minijerrycan": "Jerrybag",
    "driedfruits(b)": "SDC",
    "jerrycanwaterbottle": "Jerrybag",
    "naturebag": "SDC",
    "maninblackwaistbag": "Jerrybag",
    "headband": "SDC",
    "presentpouch(l)": "SDC",
    "4africalioncrochet": "4Africa",
    "leafkeyring": "SDC",
    "fruitcoaster": "SDC",
    "apoyoanimal(s)": "Apoyo",
    "apoyoanimal(s)+keyring": "Apoyo",
    "phonecasecard": "SDC",
    "reflector": "Jerrybag",
    "twopocketspouch": "SDC",
    "fortunefish(l)": "SDC",
    "apoyocupcoaster": "Apoyo",
    "fortunefish(s)": "SDC",
    "kitengereflector": "SDC",
    "apoyocoinpouch": "Apoyo",
    "plasticfish": "others",
    "basicpouch": "SDC",
    "4africanewmonkey": "4Africa",
    "seahorse(l)": "SDC",
    "shoppersbag": "SDC",
    "apoyolaptoppouch(m)": "Apoyo",
    "standpouch": "SDC",
    "ribbonhairband(l)": "SDC",
    "hairband": "SDC",
    "4africagorilla": "4Africa",
    "eydutextbook": "SDC",
    "dolphin(s)": "SDC",
    "turtlexl": "SDC",
    "mulapouch": "SDC",
    "shoppersbagv3": "SDC",
    "apoyosportsbag(m)": "Apoyo",
    "apoyosportsbag(s)": "Apoyo",
    "upcycledtotebag": "SDC",
    "potgloves": "SDC",
    "originalbag(s)": "SDC",
    "pencilcasev1": "SDC",
    "gloriabag(l)": "SDC",
    "whale(l)": "SDC",
    "jerrybottle": "Jerrybag",
    "apoyocardpouch": "Apoyo",
    "apoyolaptoppouch(l)": "Apoyo",
    "wristpouch": "SDC",
    "bridgetcoinpouch": "SDC",
    "4africamaxibunny": "4Africa",
    "elephantbag(l)": "SDC",
    "totebag": "SDC",
    "wristbag(l)": "SDC",
    "k.luckybagkitenge": "SDC",
    "classicwallet": "SDC",
    "ropeheadband": "SDC",
    "summerhat": "SDC",
    "presentpouch(s)": "SDC",
    "yakanacreatives1": "SDC",
    "seahorse(s)": "SDC",
    "apoyolaptoppouch(s)": "Apoyo",
    "dolphin(l)": "SDC",
    "elephantbag(s)": "SDC",
    "jerryscarf": "SDC",
    "repairservice": "SDC",
    "shoppersbagv2": "SDC",
    "apoyopouch": "Apoyo",
    "sdcsquarecoaster": "SDC",
    "makeuppouch": "SDC",
    "upcycledshoulderbag": "SDC",
    "apoyosportsbag(l)": "Apoyo",
    "apoyopencilcase": "Apoyo",
    "cupholder": "SDC",
    "t-shirt": "SDC",
    "4africarhino": "4Africa",
    "sdcapron": "SDC",
    "4africasmallrabbit": "4Africa",
    "upcycledchessbag": "SDC",
    "starsheaturmericsoap": "STAR",
    "starsheacharcoalsoap": "STAR",
    "starelevaterefinedsheabutter": "STAR",
    "starelevatepuresheabutter": "STAR",
    "starpeanutbutter&simsim400g": "STAR",
    "starsheahoney350g": "STAR",
    "starshea&turmericoil60ml": "STAR",
    "starsheabodysmoothie": "STAR",
    "starhandcream50ml": "STAR",
    "starpurelipbalm": "STAR",
    "starcoffeescrub200g": "STAR",
    "gloriabag(s)": "SDC",
    "standardbag": "SDC",
    "4africabracelet": "4Africa",
    "apoyoecobag": "Apoyo",
    "apoyoreversableecobag": "Apoyo",
    "traveltag": "SDC",
    "wristbag(s)": "SDC",
    "puffhairband(l)": "SDC",
    "apoyoanimal(l)": "Apoyo",
    "apoyoneckpillow": "Apoyo",
    "basicshoulderbag": "SDC",
    "starpeanutbuttersimsim&salt": "STAR",
    "4africarattle": "4Africa",
    "tablematspoate": "SDC",
    "fruitcoasterpoate": "SDC",
    "fruitcoaster(pumpkin)": "SDC",
    "fruitcoaster(watermelon)": "SDC",
    "4africadollywithoutflap": "4Africa",
    "k.luckybagplain": "SDC",
    "4africagiraffe": "4Africa",
    "summervendormrk": "SDC",
    "starpeanutbutter400g": "STAR",
    "starsheahoney500g": "STAR",
    "fruitcoaster(apple)": "SDC",
    "yakanacreatives2": "SDC",
    "yakanacreatives2(l)": "SDC",
    "amadilu": "SDC",
    "t-shirtwithmoreprints": "SDC",
    "t-shirtwithoneprintextra": "SDC",
    "bookbagv2zip": "SDC",
    "men'swallet": "SDC",
    "apoyhairband(s)": "SDC",
    "turtle(xl)": "SDC"
  },
  "extraChannels": [
    {
      "key": "sawe",
      "label": "SAWE",
      "color": "#4A7C59"
    }
  ],
  "customOrders": [
    {
      "id": 1,
      "groupId": 1,
      "month": 1,
      "customer": "Korea Embassy",
      "amount": 1200000,
      "note": "에코백 100개 주문",
      "category": "Government Office",
      "product": "shoppers bag",
      "qty": 100
    },
    {
      "id": 2,
      "groupId": 2,
      "month": 2,
      "customer": "Fontis",
      "amount": 1350000,
      "note": "",
      "category": "Hotel/Restaurant",
      "product": "coaster",
      "qty": 360
    },
    {
      "id": 3,
      "groupId": 3,
      "month": 3,
      "customer": "NEVO",
      "amount": 770500,
      "note": "set",
      "category": "Hotel/Restaurant",
      "product": "uniform",
      "qty": 14
    },
    {
      "id": 4,
      "groupId": 4,
      "month": 3,
      "customer": "Yamasen",
      "amount": 1800000,
      "note": "",
      "category": "Hotel/Restaurant",
      "product": "Apoyo animal keyring",
      "qty": 150
    },
    {
      "id": 5,
      "groupId": 5,
      "month": 4,
      "customer": "Ambrosoli",
      "amount": 600000,
      "note": "학교",
      "category": "NGO/Institution",
      "product": "elephant bag",
      "qty": 6
    },
    {
      "id": 6,
      "groupId": 6,
      "month": 4,
      "customer": "Yamasen",
      "amount": 309000,
      "note": "",
      "category": "Hotel/Restaurant",
      "product": "coaster",
      "qty": 91
    },
    {
      "id": 7,
      "groupId": 6,
      "month": 4,
      "customer": "Yamasen",
      "amount": 0,
      "note": "",
      "category": "Hotel/Restaurant",
      "product": "mula pouch",
      "qty": 21
    },
    {
      "id": 8,
      "groupId": 8,
      "month": 4,
      "customer": "YOU ARE",
      "amount": 1233000,
      "note": "",
      "category": "NGO/Institution",
      "product": "t-shirts",
      "qty": 43
    },
    {
      "id": 9,
      "groupId": 9,
      "month": 4,
      "customer": "Aitenga",
      "amount": 2605500,
      "note": "",
      "category": "NGO/Institution",
      "product": "others",
      "qty": null
    },
    {
      "id": 10,
      "groupId": 10,
      "month": 4,
      "customer": "Fontis",
      "amount": 768000,
      "note": "",
      "category": "Hotel/Restaurant",
      "product": "napkin",
      "qty": 60
    },
    {
      "id": 11,
      "groupId": 11,
      "month": 5,
      "customer": "Fontis",
      "amount": 768000,
      "note": "Red2",
      "category": "Hotel/Restaurant",
      "product": "napkin",
      "qty": 60
    },
    {
      "id": 12,
      "groupId": 12,
      "month": 5,
      "customer": "Mr. Kim",
      "amount": 1275000,
      "note": "현금 수령 후 은행 입금",
      "category": "Individual",
      "product": "shoppers bag",
      "qty": 100
    },
    {
      "id": 13,
      "groupId": 13,
      "month": 5,
      "customer": "Finn Church",
      "amount": 1394960,
      "note": "7기 교육비",
      "category": "NGO/Institution",
      "product": "others",
      "qty": null
    },
    {
      "id": 14,
      "groupId": 14,
      "month": 5,
      "customer": "Ms. Kim E",
      "amount": 2354500,
      "note": "현금 수령 후 은행 입금",
      "category": "Individual",
      "product": "mula pouch",
      "qty": 160
    },
    {
      "id": 15,
      "groupId": 15,
      "month": 5,
      "customer": "Fontis",
      "amount": 768000,
      "note": "Beige1",
      "category": "Hotel/Restaurant",
      "product": "napkin",
      "qty": 60
    },
    {
      "id": 16,
      "groupId": 16,
      "month": 6,
      "customer": "AIM",
      "amount": 1500000,
      "note": "학교",
      "category": "NGO/Institution",
      "product": "shoppers bag",
      "qty": 150
    },
    {
      "id": 17,
      "groupId": 17,
      "month": 6,
      "customer": "Korea Embassy",
      "amount": 994500,
      "note": "",
      "category": "Government Office",
      "product": "chef coat, etc.",
      "qty": null
    },
    {
      "id": 18,
      "groupId": 17,
      "month": 6,
      "customer": "K. Embassy staff",
      "amount": 78750,
      "note": "",
      "category": "Individual",
      "product": "apron",
      "qty": null
    },
    {
      "id": 27,
      "groupId": 17,
      "month": 6,
      "customer": "K. Embassy staff",
      "amount": 0,
      "note": "for chef 4, SDC long apron 2",
      "category": "Individual",
      "product": "apron",
      "qty": 6
    },
    {
      "id": 28,
      "groupId": 17,
      "month": 6,
      "customer": "K. Embassy staff",
      "amount": 0,
      "note": "uniform, hat, apron set",
      "category": "Individual",
      "product": "uniform",
      "qty": 4
    },
    {
      "id": 19,
      "groupId": 19,
      "month": 7,
      "customer": "K. Embassy staff",
      "amount": 60000,
      "note": "",
      "category": "Individual",
      "product": "table mat",
      "qty": 2
    },
    {
      "id": 20,
      "groupId": 20,
      "month": 7,
      "customer": "K-abic",
      "amount": 600000,
      "note": "",
      "category": "NGO/Institution",
      "product": "DTF logo print",
      "qty": 40
    },
    {
      "id": 21,
      "groupId": 21,
      "month": 7,
      "customer": "KOICA Kenya",
      "amount": 1340500,
      "note": "US $383, $1=3,500 적용",
      "category": "Government Office",
      "product": "shoppers bag",
      "qty": 60
    },
    {
      "id": 22,
      "groupId": 21,
      "month": 7,
      "customer": "KOICA Kenya",
      "amount": 0,
      "note": "US $383, $1=3,500 적용",
      "category": "Government Office",
      "product": "fortune fish",
      "qty": 60
    },
    {
      "id": 23,
      "groupId": 23,
      "month": 7,
      "customer": "K. Embassy staff",
      "amount": 90000,
      "note": "",
      "category": "Individual",
      "product": "table mat",
      "qty": 6
    },
    {
      "id": 24,
      "groupId": 24,
      "month": 7,
      "customer": "K. Embassy staff",
      "amount": 90000,
      "note": "",
      "category": "Individual",
      "product": "apron",
      "qty": 5
    },
    {
      "id": 25,
      "groupId": 25,
      "month": 7,
      "customer": "K. Embassy staff",
      "amount": 168000,
      "note": "",
      "category": "Individual",
      "product": "table mat",
      "qty": 8
    },
    {
      "id": 26,
      "groupId": 25,
      "month": 7,
      "customer": "K. Embassy staff",
      "amount": 0,
      "note": "",
      "category": "Individual",
      "product": "cup coaster",
      "qty": 8
    }
  ],
  "customLeads": [
    "Father school: v-neck t-shirts",
    "Good Neighbors Kenya: shoppers bag+fortune fish",
    "Canaan: collar t-shirts",
    "Coffee at last: collar t-shirts 9장",
    "Missionary Mr. Kim: collar t-shirts+takwondobok",
    "Coffee at last: round neck t-shirts / apron"
  ],
  "dowStats": [
    {
      "day": "Mon",
      "visits": 74,
      "contacts": 28,
      "revenue": 2791650
    },
    {
      "day": "Tue",
      "visits": 116,
      "contacts": 35,
      "revenue": 4225100
    },
    {
      "day": "Wed",
      "visits": 83,
      "contacts": 26,
      "revenue": 4833500
    },
    {
      "day": "Thu",
      "visits": 70,
      "contacts": 21,
      "revenue": 4141000
    },
    {
      "day": "Fri",
      "visits": 93,
      "contacts": 29,
      "revenue": 6124500
    },
    {
      "day": "Sat",
      "visits": 293,
      "contacts": 77,
      "revenue": 8507800
    },
    {
      "day": "Sun",
      "visits": 89,
      "contacts": 26,
      "revenue": 4191700
    }
  ],
  "visitorOrigin": [
    {
      "origin": "Uganda",
      "count": 80
    },
    {
      "origin": "Korea",
      "count": 69
    },
    {
      "origin": "Germany",
      "count": 16
    },
    {
      "origin": "USA",
      "count": 15
    },
    {
      "origin": "Italy",
      "count": 9
    },
    {
      "origin": "France",
      "count": 8
    },
    {
      "origin": "China",
      "count": 7
    },
    {
      "origin": "Japan",
      "count": 5
    },
    {
      "origin": "Other",
      "count": 22
    }
  ],
  "visitorType": [
    {
      "type": "Returning / Loyal Customers",
      "count": 68,
      "color": "#2F6E68"
    },
    {
      "type": "New / First-time Visitors",
      "count": 16,
      "color": "#D69A2D"
    }
  ],
  "eventKeywords": [
    {
      "keyword": "market",
      "label": "Market/Event",
      "count": 43
    },
    {
      "keyword": "training",
      "label": "Training Program",
      "count": 22
    },
    {
      "keyword": "church",
      "label": "Church",
      "count": 10
    },
    {
      "keyword": "koica",
      "label": "KOICA",
      "count": 9
    },
    {
      "keyword": "recommend(ed)? by|referr?(ed|al)|friend of",
      "label": "Referral",
      "count": 7
    }
  ],
  "topProducts": [
    {
      "name": "Apoyo Animal (S) + Keyring",
      "qty": 190,
      "amount": 2850000
    },
    {
      "name": "Casual bag",
      "qty": 18,
      "amount": 1440000
    },
    {
      "name": "Kitenge Reflector",
      "qty": 89,
      "amount": 1335000
    },
    {
      "name": "Mula Pouch",
      "qty": 96,
      "amount": 1152000
    },
    {
      "name": "Summer Vendor Mrk",
      "qty": 23,
      "amount": 1150000
    },
    {
      "name": "Pot Gloves",
      "qty": 104,
      "amount": 1040000
    },
    {
      "name": "T-shirt",
      "qty": 51,
      "amount": 955000
    },
    {
      "name": "Book bag V2",
      "qty": 36,
      "amount": 900000
    },
    {
      "name": "Apoyo Cup Coaster",
      "qty": 148,
      "amount": 888000
    },
    {
      "name": "Wrist Bag (L)",
      "qty": 35,
      "amount": 875000
    }
  ],
  "paymentStats": [
    {
      "method": "Cash",
      "amount": 15758899.969999999
    },
    {
      "method": "Card",
      "amount": 12346300.504
    },
    {
      "method": "MTN Mobile Money",
      "amount": 5663550
    },
    {
      "method": "Bank",
      "amount": 633000
    },
    {
      "method": "Airtel Mobile Money",
      "amount": 416500
    }
  ],
  "vendorStats": [
    {
      "vendor": "SDC",
      "qty": 1433,
      "gross": 25480500,
      "net": 23456349.490000002,
      "discount": 2024150.51,
      "count": 690
    },
    {
      "vendor": "Apoyo",
      "qty": 615,
      "gross": 8321000,
      "net": 7993800.984,
      "discount": 327199.016,
      "count": 253
    },
    {
      "vendor": "4Africa",
      "qty": 40,
      "gross": 1369000,
      "net": 1337500,
      "discount": 31500,
      "count": 31
    },
    {
      "vendor": "Jerrybag",
      "qty": 91,
      "gross": 1114000,
      "net": 1075000,
      "discount": 39000,
      "count": 49
    },
    {
      "vendor": "STAR",
      "qty": 54,
      "gross": 891000,
      "net": 886600,
      "discount": 4400,
      "count": 40
    },
    {
      "vendor": "others",
      "qty": 23,
      "gross": 69000,
      "net": 69000,
      "discount": 0,
      "count": 13
    }
  ],
  "vendorMonthly": {
    "4Africa": [
      280000,
      100000,
      135000,
      186000,
      40000,
      67000,
      202500,
      154000,
      173000,
      0,
      0,
      0
    ],
    "SDC": [
      801499.99,
      1504799.5,
      4406200,
      1930700,
      2947000,
      3157000,
      1525500,
      5703850,
      1479800,
      0,
      0,
      0
    ],
    "Apoyo": [
      212000,
      746000.48,
      1864700.48,
      1192100.024,
      1207800,
      831600,
      343000,
      1163300,
      433300,
      0,
      0,
      0
    ],
    "Jerrybag": [
      225000,
      41000,
      165000,
      78000,
      297000,
      39000,
      171000,
      44000,
      15000,
      0,
      0,
      0
    ],
    "others": [
      3000,
      3000,
      6000,
      9000,
      3000,
      18000,
      9000,
      6000,
      12000,
      0,
      0,
      0
    ],
    "STAR": [
      0,
      0,
      0,
      0,
      409500,
      182500,
      294600,
      0,
      0,
      0,
      0,
      0
    ],
    "Yakana": [
      0,
      0,
      0,
      0,
      0,
      0,
      120000,
      0,
      0,
      0,
      0,
      0
    ]
  },
  "posInvoiceCounts": [
    27,
    33,
    88,
    45,
    42,
    62,
    55,
    77,
    20,
    0,
    0,
    0
  ],
  "posMonthlyQty": [
    77,
    186,
    440,
    297,
    303,
    354,
    106,
    391,
    102,
    0,
    0,
    0
  ],
  "allProducts": [
    {
      "name": "Apoyo Animal (L)",
      "qty": [
        0,
        0,
        0,
        0,
        2,
        0,
        0,
        0,
        4,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        50000,
        0,
        0,
        0,
        100000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Animal (S)",
      "qty": [
        1,
        0,
        6,
        2,
        0,
        1,
        7,
        2,
        0,
        0,
        0,
        0
      ],
      "amount": [
        13000,
        0,
        78000,
        26000,
        0,
        13000,
        91000,
        26000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Animal (S) + Keyring",
      "qty": [
        9,
        11,
        57,
        27,
        26,
        12,
        8,
        31,
        9,
        0,
        0,
        0
      ],
      "amount": [
        135000,
        165000,
        855000,
        405000,
        390000,
        180000,
        120000,
        465000,
        135000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Neck Pillow",
      "qty": [
        0,
        0,
        0,
        0,
        1,
        0,
        0,
        0,
        1,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        31000,
        0,
        0,
        0,
        31000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Hair Band (S)",
      "qty": [
        2,
        6,
        8,
        11,
        2,
        9,
        8,
        12,
        3,
        0,
        0,
        0
      ],
      "amount": [
        6000,
        18000,
        24000,
        33000,
        6000,
        27000,
        24000,
        36000,
        9000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Hair Band (L)",
      "qty": [
        2,
        12,
        5,
        7,
        1,
        3,
        2,
        1,
        1,
        0,
        0,
        0
      ],
      "amount": [
        10000,
        60000,
        25000,
        35000,
        5000,
        15000,
        10000,
        5000,
        5000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Cup Coaster",
      "qty": [
        6,
        12,
        31,
        42,
        3,
        24,
        0,
        30,
        0,
        0,
        0,
        0
      ],
      "amount": [
        36000,
        72000,
        186000,
        252000,
        18000,
        144000,
        0,
        180000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Pouch",
      "qty": [
        0,
        0,
        3,
        9,
        5,
        5,
        0,
        6,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        45000,
        135000,
        75000,
        75000,
        0,
        90000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Card Pouch",
      "qty": [
        0,
        1,
        1,
        5,
        5,
        4,
        0,
        4,
        4,
        0,
        0,
        0
      ],
      "amount": [
        0,
        20000,
        20000,
        100000,
        100000,
        80000,
        0,
        80000,
        80000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Coin Pouch",
      "qty": [
        1,
        15,
        2,
        3,
        7,
        1,
        0,
        4,
        0,
        0,
        0,
        0
      ],
      "amount": [
        12000,
        180000,
        24000,
        36000,
        84000,
        12000,
        0,
        48000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Pencil Case",
      "qty": [
        0,
        0,
        0,
        4,
        2,
        4,
        0,
        2,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        60000,
        30000,
        60000,
        0,
        30000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Eco Bag",
      "qty": [
        0,
        0,
        0,
        0,
        3,
        1,
        0,
        3,
        2,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        126000,
        42000,
        0,
        126000,
        84000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Reversable Eco Bag",
      "qty": [
        0,
        0,
        0,
        0,
        2,
        1,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        84000,
        42000,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Laptop Pouch (L)",
      "qty": [
        0,
        1,
        3,
        2,
        2,
        2,
        1,
        2,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        52000,
        156000,
        104000,
        104000,
        104000,
        52000,
        104000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Laptop Pouch (M)",
      "qty": [
        0,
        3,
        5,
        1,
        1,
        1,
        1,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        138000,
        230000,
        46000,
        46000,
        46000,
        46000,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Laptop Pouch (S)",
      "qty": [
        0,
        0,
        2,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        80000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Sports Bag (L)",
      "qty": [
        0,
        0,
        1,
        2,
        1,
        0,
        0,
        1,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        31000,
        62000,
        31000,
        0,
        0,
        31000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Sports Bag (M)",
      "qty": [
        0,
        1,
        2,
        1,
        1,
        1,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        28000,
        56000,
        28000,
        28000,
        28000,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Apoyo Sports Bag (S)",
      "qty": [
        0,
        1,
        4,
        0,
        1,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        25000,
        100000,
        0,
        25000,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Big Elephant",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa New Elephant",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Big Hippo",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa New Monkey",
      "qty": [
        0,
        1,
        1,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        70000,
        70000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Big Monkey",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Crocodile",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Snake",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Rattle",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        1,
        1,
        1,
        1,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        25000,
        25000,
        25000,
        25000,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Maxi Bunny",
      "qty": [
        0,
        0,
        1,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        50000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Small Rabbit",
      "qty": [
        0,
        0,
        0,
        1,
        0,
        0,
        0,
        0,
        1,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        50000,
        0,
        0,
        0,
        0,
        50000,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Boho with straps",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Kids Apron",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Lion Crochet",
      "qty": [
        1,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        1,
        0,
        0,
        0
      ],
      "amount": [
        70000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        70000,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Zebra",
      "qty": [
        3,
        0,
        0,
        0,
        0,
        0,
        0,
        1,
        0,
        0,
        0,
        0
      ],
      "amount": [
        210000,
        0,
        0,
        0,
        0,
        0,
        0,
        70000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Rhino",
      "qty": [
        0,
        0,
        0,
        1,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        70000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Gorilla",
      "qty": [
        0,
        2,
        1,
        5,
        0,
        3,
        0,
        2,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        30000,
        15000,
        75000,
        0,
        60000,
        0,
        40000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Bracelet",
      "qty": [
        0,
        0,
        0,
        0,
        2,
        0,
        0,
        3,
        4,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        40000,
        0,
        0,
        21000,
        28000,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Macreme Bag",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Giraffe",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        1,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        80000,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "4Africa Dolly without Flap",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        1,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        100000,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Amaro Morning Flakes",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Coffee Scrub 200g",
      "qty": [
        0,
        0,
        0,
        0,
        3,
        1,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        75000,
        25000,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Hand Cream 50ml",
      "qty": [
        0,
        0,
        0,
        0,
        1,
        0,
        2,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        19000,
        0,
        38000,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Peanut Butter & Sim sim 400g",
      "qty": [
        0,
        0,
        0,
        0,
        1,
        0,
        2,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        7500,
        0,
        15000,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Peanut Butter 400g",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        1,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        7500,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Peanut Butter Sim sim & Salt",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        2,
        2,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        8000,
        8000,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Pure Lip Balm",
      "qty": [
        0,
        0,
        0,
        0,
        1,
        2,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        23000,
        46000,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Shea Honey 350g",
      "qty": [
        0,
        0,
        0,
        0,
        3,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        44000,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Shea Honey 500g",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        1,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        22000,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Shea Honey 1kg",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Shea & Turmeric oil 60ml",
      "qty": [
        0,
        0,
        0,
        0,
        3,
        0,
        2,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        66000,
        0,
        44000,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Shea & Turmeric oil 100ml",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Shea & Turmeric oil 250ml",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Elevate Pure Shea Butter",
      "qty": [
        0,
        0,
        0,
        0,
        1,
        1,
        2,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        22000,
        22000,
        44000,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Elevate Refined Shea Butter",
      "qty": [
        0,
        0,
        0,
        0,
        3,
        10,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        21000,
        70000,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Shea Turmeric Soap",
      "qty": [
        0,
        0,
        0,
        0,
        2,
        1,
        2,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        28000,
        14000,
        28000,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Shea Charcoal Soap",
      "qty": [
        0,
        0,
        0,
        0,
        1,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        14000,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "STAR Shea Body Smoothie",
      "qty": [
        0,
        0,
        0,
        0,
        2,
        0,
        2,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        90000,
        0,
        90000,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Art Collabo Angel On Earth",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Amadilu",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        1,
        1,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        45000,
        45000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Basic Shoulder bag",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        2,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        60000,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Book bag V2 Zip",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        1,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        30000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Basic Pouch",
      "qty": [
        2,
        2,
        6,
        5,
        0,
        0,
        0,
        3,
        1,
        0,
        0,
        0
      ],
      "amount": [
        6000,
        6000,
        18000,
        15000,
        0,
        0,
        0,
        30000,
        10000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Book bag (discontinued)",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Book bag V2",
      "qty": [
        1,
        0,
        6,
        5,
        9,
        7,
        0,
        6,
        2,
        0,
        0,
        0
      ],
      "amount": [
        25000,
        0,
        150000,
        125000,
        225000,
        175000,
        0,
        150000,
        50000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Bottle Bag (discontinued)",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Bottle Sticker",
      "qty": [
        1,
        0,
        3,
        6,
        1,
        4,
        5,
        3,
        0,
        0,
        0,
        0
      ],
      "amount": [
        2500,
        0,
        7500,
        15000,
        2500,
        10000,
        12500,
        7500,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Bottle Ring (discontinued)",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Bow Tie (discontinued)",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Bridget Coin Pouch",
      "qty": [
        0,
        0,
        1,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        5000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Casual bag",
      "qty": [
        3,
        2,
        2,
        2,
        2,
        2,
        0,
        4,
        1,
        0,
        0,
        0
      ],
      "amount": [
        240000,
        160000,
        160000,
        160000,
        160000,
        160000,
        0,
        320000,
        80000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Casual extra bag (ODM only)",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Classic Wallet",
      "qty": [
        0,
        0,
        3,
        0,
        0,
        3,
        1,
        3,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        75000,
        0,
        0,
        75000,
        25000,
        75000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Cup holder",
      "qty": [
        0,
        0,
        0,
        8,
        0,
        2,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        80000,
        0,
        20000,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Cushion (discontinued)",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Design Tag (discontinued)",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Dolphin (L)",
      "qty": [
        0,
        0,
        1,
        2,
        0,
        0,
        1,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        20000,
        40000,
        0,
        0,
        20000,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Dolphin (S)",
      "qty": [
        0,
        1,
        2,
        0,
        1,
        2,
        0,
        0,
        2,
        0,
        0,
        0
      ],
      "amount": [
        0,
        15000,
        30000,
        0,
        15000,
        30000,
        0,
        0,
        30000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Eg Bag (discontinued)",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Elephant bag (L)",
      "qty": [
        0,
        0,
        1,
        0,
        0,
        0,
        0,
        4,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        120000,
        0,
        0,
        0,
        0,
        480000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Elephant bag (S)",
      "qty": [
        0,
        0,
        1,
        0,
        1,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        100000,
        0,
        100000,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "EYDU Textbook",
      "qty": [
        0,
        1,
        1,
        0,
        0,
        0,
        0,
        1,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        25000,
        35000,
        0,
        0,
        0,
        0,
        25000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Flower keyring",
      "qty": [
        1,
        53,
        8,
        15,
        3,
        6,
        2,
        19,
        4,
        0,
        0,
        0
      ],
      "amount": [
        7000,
        371000,
        56000,
        105000,
        21000,
        42000,
        14000,
        133000,
        28000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Fortune Fish (L)",
      "qty": [
        1,
        0,
        19,
        13,
        0,
        4,
        2,
        4,
        2,
        0,
        0,
        0
      ],
      "amount": [
        15000,
        0,
        285000,
        195000,
        0,
        60000,
        30000,
        60000,
        30000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Fortune Fish (S)",
      "qty": [
        1,
        0,
        28,
        5,
        11,
        13,
        1,
        4,
        1,
        0,
        0,
        0
      ],
      "amount": [
        10000,
        0,
        280000,
        50000,
        110000,
        130000,
        10000,
        40000,
        10000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Fruit Coaster",
      "qty": [
        8,
        0,
        7,
        1,
        0,
        4,
        0,
        4,
        6,
        0,
        0,
        0
      ],
      "amount": [
        80000,
        0,
        70000,
        10000,
        0,
        40000,
        0,
        40000,
        60000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Gloria bag (L)",
      "qty": [
        0,
        1,
        0,
        0,
        0,
        0,
        0,
        4,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        150000,
        0,
        0,
        0,
        0,
        0,
        600000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Gloria bag (S)",
      "qty": [
        0,
        0,
        0,
        0,
        1,
        0,
        0,
        1,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        100000,
        0,
        0,
        100000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Hair Band",
      "qty": [
        0,
        4,
        5,
        3,
        8,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        12000,
        15000,
        9000,
        24000,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Ribbon Hair Band (L)",
      "qty": [
        0,
        1,
        0,
        1,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        15000,
        0,
        15000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Ribbon Hair Band (S)",
      "qty": [
        1,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        7000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Repair Service",
      "qty": [
        0,
        0,
        2,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        34000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Head band",
      "qty": [
        1,
        1,
        6,
        0,
        1,
        4,
        0,
        2,
        1,
        0,
        0,
        0
      ],
      "amount": [
        13000,
        13000,
        78000,
        0,
        13000,
        52000,
        0,
        26000,
        13000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Jerry Calender",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Jerry Handkerchief",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Jerry Rain Cover",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Jerry Scarf",
      "qty": [
        0,
        0,
        4,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        12000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Jerry Bottle",
      "qty": [
        0,
        1,
        0,
        0,
        0,
        0,
        0,
        1,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        20000,
        0,
        0,
        0,
        0,
        0,
        20000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Jerrycan Water Bottle",
      "qty": [
        2,
        1,
        6,
        0,
        18,
        0,
        0,
        0,
        1,
        0,
        0,
        0
      ],
      "amount": [
        30000,
        15000,
        90000,
        0,
        270000,
        0,
        0,
        0,
        15000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Jerrycan bag",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Key holder",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "K. Lucky bag kitenge",
      "qty": [
        0,
        0,
        1,
        0,
        0,
        1,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        60000,
        0,
        0,
        60000,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "K. Lucky bag plain",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        1,
        1,
        1,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        80000,
        80000,
        80000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Leaf keyring",
      "qty": [
        5,
        2,
        5,
        0,
        0,
        1,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        25000,
        10000,
        25000,
        0,
        0,
        5000,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Logo",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Make Up Pouch",
      "qty": [
        0,
        0,
        1,
        0,
        5,
        4,
        0,
        2,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        20000,
        0,
        100000,
        70000,
        0,
        30000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Man In Black Waist Bag",
      "qty": [
        1,
        0,
        0,
        0,
        0,
        0,
        1,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        180000,
        0,
        0,
        0,
        0,
        0,
        180000,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Men's Wallet",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        1,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        25000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Mini Jerrycan",
      "qty": [
        1,
        2,
        9,
        2,
        16,
        5,
        3,
        8,
        0,
        0,
        0,
        0
      ],
      "amount": [
        3000,
        6000,
        27000,
        6000,
        48000,
        15000,
        9000,
        24000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Nature Bag",
      "qty": [
        1,
        0,
        1,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        80000,
        0,
        80000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Original bag (L)",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Original bag (S)",
      "qty": [
        0,
        1,
        1,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        170000,
        170000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Pajama",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Paper bag L",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Paper bag M",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Paper bag S",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Pencil Case V1",
      "qty": [
        0,
        1,
        0,
        7,
        5,
        1,
        0,
        11,
        1,
        0,
        0,
        0
      ],
      "amount": [
        0,
        10000,
        0,
        70000,
        50000,
        10000,
        0,
        110000,
        10000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Picnic bag",
      "qty": [
        1,
        0,
        6,
        0,
        5,
        8,
        0,
        3,
        1,
        0,
        0,
        0
      ],
      "amount": [
        30000,
        0,
        180000,
        0,
        150000,
        240000,
        0,
        90000,
        30000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Plastic Fish",
      "qty": [
        1,
        1,
        2,
        3,
        1,
        6,
        3,
        2,
        4,
        0,
        0,
        0
      ],
      "amount": [
        3000,
        3000,
        6000,
        9000,
        3000,
        18000,
        9000,
        6000,
        12000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Pot Gloves",
      "qty": [
        0,
        5,
        13,
        20,
        22,
        20,
        0,
        24,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        50000,
        130000,
        200000,
        220000,
        200000,
        0,
        240000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Pot Mat (L) (discontinued)",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Pot Mat (S) (discontinued)",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Present pouch (L)",
      "qty": [
        2,
        0,
        4,
        2,
        0,
        8,
        2,
        6,
        0,
        0,
        0,
        0
      ],
      "amount": [
        30000,
        0,
        60000,
        30000,
        0,
        120000,
        30000,
        90000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Present pouch (S)",
      "qty": [
        0,
        0,
        5,
        3,
        6,
        7,
        0,
        6,
        1,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        50000,
        30000,
        60000,
        70000,
        0,
        60000,
        10000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Phone Case Card",
      "qty": [
        1,
        0,
        10,
        6,
        2,
        10,
        1,
        9,
        1,
        0,
        0,
        0
      ],
      "amount": [
        3000,
        0,
        30000,
        18000,
        6000,
        30000,
        3000,
        27000,
        3000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Puff Hair Band (S)",
      "qty": [
        3,
        2,
        6,
        1,
        3,
        0,
        0,
        5,
        1,
        0,
        0,
        0
      ],
      "amount": [
        15000,
        10000,
        30000,
        5000,
        15000,
        0,
        0,
        25000,
        5000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Puff Hair Band (L)",
      "qty": [
        0,
        0,
        0,
        0,
        2,
        1,
        0,
        2,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        16000,
        8000,
        0,
        16000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Reflector",
      "qty": [
        1,
        0,
        4,
        6,
        0,
        2,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        12000,
        0,
        48000,
        72000,
        0,
        24000,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Kitenge Reflector",
      "qty": [
        1,
        1,
        42,
        2,
        12,
        9,
        1,
        19,
        2,
        0,
        0,
        0
      ],
      "amount": [
        15000,
        15000,
        630000,
        30000,
        180000,
        135000,
        15000,
        285000,
        30000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Rope Head Band",
      "qty": [
        0,
        0,
        2,
        0,
        0,
        0,
        0,
        8,
        4,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        50000,
        0,
        0,
        0,
        0,
        200000,
        100000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Sample_Bag",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Sample_Flower & leaf keyring",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Sample_pouch",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Seahorse (L)",
      "qty": [
        0,
        4,
        3,
        0,
        3,
        0,
        0,
        1,
        2,
        0,
        0,
        0
      ],
      "amount": [
        0,
        60000,
        45000,
        0,
        45000,
        0,
        0,
        15000,
        30000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Seahorse (S)",
      "qty": [
        0,
        0,
        12,
        11,
        5,
        3,
        0,
        2,
        1,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        120000,
        110000,
        50000,
        30000,
        0,
        20000,
        10000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Shoppers bag",
      "qty": [
        0,
        1,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        15000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Shoppers bag V2",
      "qty": [
        0,
        0,
        1,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        10000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Shoppers bag V3",
      "qty": [
        0,
        1,
        3,
        4,
        4,
        2,
        2,
        4,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        15000,
        45000,
        60000,
        60000,
        30000,
        30000,
        60000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Mula Pouch",
      "qty": [
        0,
        19,
        2,
        1,
        18,
        34,
        0,
        20,
        2,
        0,
        0,
        0
      ],
      "amount": [
        0,
        228000,
        24000,
        12000,
        216000,
        408000,
        0,
        240000,
        24000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Stand Pouch",
      "qty": [
        0,
        1,
        16,
        6,
        1,
        6,
        1,
        4,
        2,
        0,
        0,
        0
      ],
      "amount": [
        0,
        3000,
        48000,
        18000,
        3000,
        46000,
        10000,
        40000,
        20000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Standard bag",
      "qty": [
        0,
        0,
        0,
        0,
        1,
        2,
        0,
        0,
        1,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        80000,
        160000,
        0,
        0,
        80000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Summer Vendor Mrk",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        15,
        8,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        750000,
        400000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Summer Hat",
      "qty": [
        0,
        0,
        3,
        0,
        0,
        0,
        0,
        1,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        45000,
        0,
        0,
        0,
        0,
        15000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "SDC Square Coaster",
      "qty": [
        0,
        0,
        2,
        5,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        12000,
        30000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "SDC Apron",
      "qty": [
        0,
        0,
        0,
        3,
        3,
        3,
        0,
        5,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        105000,
        105000,
        105000,
        0,
        175000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Fruit Coaster Poate",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        4,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        40000,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Table Mats Poate",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        4,
        0,
        4,
        2,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        120000,
        0,
        120000,
        60000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Tyde's 1",
      "qty": [
        1,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        95000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Tyde's 2",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Tyde's 3",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Table mat",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Tote bag",
      "qty": [
        0,
        0,
        1,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        60000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Travel Tag",
      "qty": [
        0,
        0,
        0,
        0,
        2,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        26000,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Two Pockets Pouch",
      "qty": [
        1,
        1,
        6,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        25000,
        25000,
        150000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Turtle (S)",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Turtle (L)",
      "qty": [
        3,
        1,
        4,
        3,
        1,
        4,
        2,
        0,
        7,
        0,
        0,
        0
      ],
      "amount": [
        90000,
        30000,
        120000,
        90000,
        30000,
        120000,
        60000,
        0,
        210000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Turtle XL",
      "qty": [
        0,
        1,
        2,
        0,
        1,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        100000,
        200000,
        0,
        100000,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "T- shirt with one print extra",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        2,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        76000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "T-shirt with more prints",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        2,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        82000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "T- shirt 2 (SDC)",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "T-shirt",
      "qty": [
        0,
        0,
        0,
        2,
        2,
        13,
        7,
        27,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        80000,
        30000,
        195000,
        105000,
        545000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Yakana Creatives 1",
      "qty": [
        0,
        0,
        1,
        1,
        3,
        0,
        2,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        60000,
        60000,
        180000,
        0,
        120000,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Yakana Creatives 2 (S)",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Yakana Creatives 2 (L)",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        1,
        1,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        40000,
        40000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Yakana Creatives 3",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Upcycled Tote bag",
      "qty": [
        0,
        1,
        0,
        0,
        0,
        0,
        0,
        2,
        1,
        0,
        0,
        0
      ],
      "amount": [
        0,
        80000,
        0,
        0,
        0,
        0,
        0,
        160000,
        80000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Upcycled Shoulder bag",
      "qty": [
        0,
        0,
        1,
        0,
        0,
        0,
        1,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        150000,
        0,
        0,
        0,
        150000,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Upcycled Chess bag",
      "qty": [
        0,
        0,
        0,
        2,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        300000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Dried Fruits (B)",
      "qty": [
        2,
        0,
        1,
        6,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        26000,
        0,
        13000,
        52000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Dried Fruits (S)",
      "qty": [
        1,
        0,
        0,
        11,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        10000,
        0,
        0,
        60000,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Xmas Stockings",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Xmas deco Bell",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Xmas deco Star",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Xmas deco Tree",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Xmas deco Giftbox",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Xmas deco Ball",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Whale (L)",
      "qty": [
        0,
        1,
        8,
        1,
        4,
        0,
        0,
        1,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        20000,
        160000,
        20000,
        80000,
        0,
        0,
        20000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Whale (S)",
      "qty": [
        4,
        6,
        10,
        0,
        0,
        2,
        2,
        3,
        1,
        0,
        0,
        0
      ],
      "amount": [
        60000,
        90000,
        150000,
        0,
        0,
        30000,
        30000,
        45000,
        15000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Wrist Bag (S)",
      "qty": [
        0,
        0,
        0,
        0,
        19,
        18,
        4,
        6,
        1,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        285000,
        270000,
        60000,
        90000,
        15000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Wrist Bag (L)",
      "qty": [
        0,
        0,
        6,
        0,
        8,
        5,
        0,
        13,
        3,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        150000,
        0,
        200000,
        125000,
        0,
        325000,
        75000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Wrist pouch",
      "qty": [
        0,
        0,
        1,
        0,
        5,
        1,
        0,
        7,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        20000,
        0,
        100000,
        20000,
        0,
        140000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Fruit Coaster (Pumpkin)",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        2,
        0,
        1,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        20000,
        0,
        10000,
        0,
        0,
        0,
        0
      ]
    },
    {
      "name": "Fruit Coaster (Watermelon)",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        25,
        0,
        2,
        1,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        250000,
        0,
        20000,
        10000,
        0,
        0,
        0
      ]
    },
    {
      "name": "Workshop",
      "qty": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ],
      "amount": [
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0
      ]
    }
  ],
  "vendorMonthlyDetail": {
    "4Africa": [
      {
        "qty": 4,
        "gross": 280000,
        "net": 280000,
        "discount": 0,
        "count": 3
      },
      {
        "qty": 3,
        "gross": 100000,
        "net": 100000,
        "discount": 0,
        "count": 2
      },
      {
        "qty": 3,
        "gross": 135000,
        "net": 135000,
        "discount": 0,
        "count": 3
      },
      {
        "qty": 7,
        "gross": 195000,
        "net": 186000,
        "discount": 9000,
        "count": 4
      },
      {
        "qty": 2,
        "gross": 40000,
        "net": 40000,
        "discount": 0,
        "count": 1
      },
      {
        "qty": 4,
        "gross": 85000,
        "net": 67000,
        "discount": 18000,
        "count": 3
      },
      {
        "qty": 3,
        "gross": 205000,
        "net": 202500,
        "discount": 2500,
        "count": 3
      },
      {
        "qty": 7,
        "gross": 156000,
        "net": 154000,
        "discount": 2000,
        "count": 7
      },
      {
        "qty": 7,
        "gross": 173000,
        "net": 173000,
        "discount": 0,
        "count": 5
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      }
    ],
    "SDC": [
      {
        "qty": 46,
        "gross": 909500,
        "net": 801499.99,
        "discount": 108000.01000000001,
        "count": 29
      },
      {
        "qty": 115,
        "gross": 1698000,
        "net": 1504799.5,
        "discount": 193200.5,
        "count": 37
      },
      {
        "qty": 286,
        "gross": 4647500,
        "net": 4406200,
        "discount": 241300,
        "count": 152
      },
      {
        "qty": 163,
        "gross": 2199000,
        "net": 1930700,
        "discount": 268300,
        "count": 80
      },
      {
        "qty": 180,
        "gross": 3157500,
        "net": 2947000,
        "discount": 210500,
        "count": 72
      },
      {
        "qty": 251,
        "gross": 3771000,
        "net": 3157000,
        "discount": 614000,
        "count": 103
      },
      {
        "qty": 53,
        "gross": 1554500,
        "net": 1525500,
        "discount": 29000,
        "count": 42
      },
      {
        "qty": 275,
        "gross": 6022500,
        "net": 5703850,
        "discount": 318650,
        "count": 131
      },
      {
        "qty": 64,
        "gross": 1521000,
        "net": 1479800,
        "discount": 41200,
        "count": 44
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      }
    ],
    "Apoyo": [
      {
        "qty": 21,
        "gross": 212000,
        "net": 212000,
        "discount": 0,
        "count": 8
      },
      {
        "qty": 63,
        "gross": 758000,
        "net": 746000.48,
        "discount": 11999.52,
        "count": 19
      },
      {
        "qty": 130,
        "gross": 1910000,
        "net": 1864700.48,
        "discount": 45299.520000000004,
        "count": 55
      },
      {
        "qty": 116,
        "gross": 1322000,
        "net": 1192100.024,
        "discount": 129899.976,
        "count": 41
      },
      {
        "qty": 65,
        "gross": 1233000,
        "net": 1207800,
        "discount": 25200,
        "count": 38
      },
      {
        "qty": 69,
        "gross": 868000,
        "net": 831600,
        "discount": 36400,
        "count": 34
      },
      {
        "qty": 27,
        "gross": 343000,
        "net": 343000,
        "discount": 0,
        "count": 9
      },
      {
        "qty": 98,
        "gross": 1221000,
        "net": 1163300,
        "discount": 57700,
        "count": 32
      },
      {
        "qty": 26,
        "gross": 454000,
        "net": 433300,
        "discount": 20700,
        "count": 17
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      }
    ],
    "Jerrybag": [
      {
        "qty": 5,
        "gross": 225000,
        "net": 225000,
        "discount": 0,
        "count": 4
      },
      {
        "qty": 4,
        "gross": 41000,
        "net": 41000,
        "discount": 0,
        "count": 4
      },
      {
        "qty": 19,
        "gross": 165000,
        "net": 165000,
        "discount": 0,
        "count": 12
      },
      {
        "qty": 8,
        "gross": 78000,
        "net": 78000,
        "discount": 0,
        "count": 3
      },
      {
        "qty": 34,
        "gross": 318000,
        "net": 297000,
        "discount": 21000,
        "count": 9
      },
      {
        "qty": 7,
        "gross": 39000,
        "net": 39000,
        "discount": 0,
        "count": 5
      },
      {
        "qty": 4,
        "gross": 189000,
        "net": 171000,
        "discount": 18000,
        "count": 4
      },
      {
        "qty": 9,
        "gross": 44000,
        "net": 44000,
        "discount": 0,
        "count": 7
      },
      {
        "qty": 1,
        "gross": 15000,
        "net": 15000,
        "discount": 0,
        "count": 1
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      }
    ],
    "others": [
      {
        "qty": 1,
        "gross": 3000,
        "net": 3000,
        "discount": 0,
        "count": 1
      },
      {
        "qty": 1,
        "gross": 3000,
        "net": 3000,
        "discount": 0,
        "count": 1
      },
      {
        "qty": 2,
        "gross": 6000,
        "net": 6000,
        "discount": 0,
        "count": 1
      },
      {
        "qty": 3,
        "gross": 9000,
        "net": 9000,
        "discount": 0,
        "count": 1
      },
      {
        "qty": 1,
        "gross": 3000,
        "net": 3000,
        "discount": 0,
        "count": 1
      },
      {
        "qty": 6,
        "gross": 18000,
        "net": 18000,
        "discount": 0,
        "count": 3
      },
      {
        "qty": 3,
        "gross": 9000,
        "net": 9000,
        "discount": 0,
        "count": 2
      },
      {
        "qty": 2,
        "gross": 6000,
        "net": 6000,
        "discount": 0,
        "count": 1
      },
      {
        "qty": 4,
        "gross": 12000,
        "net": 12000,
        "discount": 0,
        "count": 2
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      }
    ],
    "STAR": [
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 21,
        "gross": 409500,
        "net": 409500,
        "discount": 0,
        "count": 18
      },
      {
        "qty": 17,
        "gross": 185000,
        "net": 182500,
        "discount": 2500,
        "count": 8
      },
      {
        "qty": 16,
        "gross": 296500,
        "net": 294600,
        "discount": 1900,
        "count": 14
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      }
    ],
    "Yakana": [
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 2,
        "gross": 120000,
        "net": 120000,
        "discount": 0,
        "count": 2
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      },
      {
        "qty": 0,
        "gross": 0,
        "net": 0,
        "discount": 0,
        "count": 0
      }
    ]
  },
  "operationStats": {
    "operatingDays": 270,
    "zeroVisitDays": 76,
    "zeroVisitPct": 0.2814814814814815,
    "zeroSalesDays": 109,
    "zeroSalesPct": 0.40370370370370373,
    "top5Concentration": 0.18735180703858223,
    "totalRevenue": 34815250
  },
  "marketEvents": [
    {
      "id": 1,
      "date": "2026-03-28",
      "month": 2,
      "name": "SDC Easter Design Market",
      "revenue": 1573300,
      "qty": 127,
      "invoices": 39
    },
    {
      "id": 2,
      "date": "2026-05-01",
      "month": 4,
      "name": "Korean Community Sports Day",
      "revenue": 623000,
      "qty": 44,
      "invoices": 2
    },
    {
      "id": 3,
      "date": "2026-05-09",
      "month": 4,
      "name": "Ambrosoli Family Fun Day",
      "revenue": 201000,
      "qty": 13,
      "invoices": 2
    },
    {
      "id": 4,
      "date": "2026-06-26",
      "month": 5,
      "name": "mini market day",
      "revenue": 1161400,
      "qty": 110,
      "invoices": 14
    },
    {
      "id": 5,
      "date": "2026-07-25",
      "month": 6,
      "name": "July mini market day",
      "revenue": 205500,
      "qty": 18,
      "invoices": 2
    },
    {
      "id": 6,
      "date": "2026-08-15",
      "month": 7,
      "name": "Heritage Market",
      "revenue": 20000,
      "qty": 2,
      "invoices": 2
    },
    {
      "id": 7,
      "date": "2026-08-29",
      "month": 7,
      "name": "SDC Summer Design Market",
      "revenue": 1212000,
      "qty": 56,
      "invoices": 19
    }
  ],
  "productPerformance": {
    "totalCount": 168,
    "zeroOrOneCount": 53,
    "lessThan5Count": 95,
    "underperformers": [
      "4Africa Big Elephant(0)",
      "4Africa New Elephant(0)",
      "4Africa Big Hippo(0)",
      "4Africa Big Monkey(0)",
      "4Africa Crocodile(0)",
      "4Africa Snake(0)",
      "4Africa Boho with straps(0)",
      "4Africa Kids Apron(0)",
      "4Africa Macreme Bag(0)",
      "STAR Amaro Morning Flakes(0)",
      "STAR Shea Honey 1kg(0)",
      "STAR Shea & Turmeric oil 100ml(0)",
      "STAR Shea & Turmeric oil 250ml(0)",
      "Art Collabo Angel On Earth(0)",
      "Casual extra bag (ODM only)(0)"
    ],
    "allSellable": [
      {
        "name": "Apoyo Animal (L)",
        "qty": 6,
        "amount": 150000
      },
      {
        "name": "Apoyo Animal (S)",
        "qty": 19,
        "amount": 247000
      },
      {
        "name": "Apoyo Animal (S) + Keyring",
        "qty": 190,
        "amount": 2850000
      },
      {
        "name": "Apoyo Neck Pillow",
        "qty": 2,
        "amount": 62000
      },
      {
        "name": "Apoyo Hair Band (S)",
        "qty": 61,
        "amount": 183000
      },
      {
        "name": "Apoyo Hair Band (L)",
        "qty": 34,
        "amount": 170000
      },
      {
        "name": "Apoyo Cup Coaster",
        "qty": 148,
        "amount": 888000
      },
      {
        "name": "Apoyo Pouch",
        "qty": 28,
        "amount": 420000
      },
      {
        "name": "Apoyo Card Pouch",
        "qty": 24,
        "amount": 480000
      },
      {
        "name": "Apoyo Coin Pouch",
        "qty": 33,
        "amount": 396000
      },
      {
        "name": "Apoyo Pencil Case",
        "qty": 12,
        "amount": 180000
      },
      {
        "name": "Apoyo Eco Bag",
        "qty": 9,
        "amount": 378000
      },
      {
        "name": "Apoyo Reversable Eco Bag",
        "qty": 3,
        "amount": 126000
      },
      {
        "name": "Apoyo Laptop Pouch (L)",
        "qty": 13,
        "amount": 676000
      },
      {
        "name": "Apoyo Laptop Pouch (M)",
        "qty": 12,
        "amount": 552000
      },
      {
        "name": "Apoyo Laptop Pouch (S)",
        "qty": 2,
        "amount": 80000
      },
      {
        "name": "Apoyo Sports Bag (L)",
        "qty": 5,
        "amount": 155000
      },
      {
        "name": "Apoyo Sports Bag (M)",
        "qty": 6,
        "amount": 168000
      },
      {
        "name": "Apoyo Sports Bag (S)",
        "qty": 6,
        "amount": 150000
      },
      {
        "name": "4Africa Big Elephant",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "4Africa New Elephant",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "4Africa Big Hippo",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "4Africa New Monkey",
        "qty": 2,
        "amount": 140000
      },
      {
        "name": "4Africa Big Monkey",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "4Africa Crocodile",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "4Africa Snake",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "4Africa Rattle",
        "qty": 4,
        "amount": 100000
      },
      {
        "name": "4Africa Maxi Bunny",
        "qty": 1,
        "amount": 50000
      },
      {
        "name": "4Africa Small Rabbit",
        "qty": 2,
        "amount": 100000
      },
      {
        "name": "4Africa Boho with straps",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "4Africa Kids Apron",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "4Africa Lion Crochet",
        "qty": 2,
        "amount": 140000
      },
      {
        "name": "4Africa Zebra",
        "qty": 4,
        "amount": 280000
      },
      {
        "name": "4Africa Rhino",
        "qty": 1,
        "amount": 70000
      },
      {
        "name": "4Africa Gorilla",
        "qty": 13,
        "amount": 220000
      },
      {
        "name": "4Africa Bracelet",
        "qty": 9,
        "amount": 89000
      },
      {
        "name": "4Africa Macreme Bag",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "4Africa Giraffe",
        "qty": 1,
        "amount": 80000
      },
      {
        "name": "4Africa Dolly without Flap",
        "qty": 1,
        "amount": 100000
      },
      {
        "name": "STAR Amaro Morning Flakes",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "STAR Coffee Scrub 200g",
        "qty": 4,
        "amount": 100000
      },
      {
        "name": "STAR Hand Cream 50ml",
        "qty": 3,
        "amount": 57000
      },
      {
        "name": "STAR Peanut Butter & Sim sim 400g",
        "qty": 3,
        "amount": 22500
      },
      {
        "name": "STAR Peanut Butter 400g",
        "qty": 1,
        "amount": 7500
      },
      {
        "name": "STAR Peanut Butter Sim sim & Salt",
        "qty": 4,
        "amount": 16000
      },
      {
        "name": "STAR Pure Lip Balm",
        "qty": 3,
        "amount": 69000
      },
      {
        "name": "STAR Shea Honey 350g",
        "qty": 3,
        "amount": 44000
      },
      {
        "name": "STAR Shea Honey 500g",
        "qty": 1,
        "amount": 22000
      },
      {
        "name": "STAR Shea Honey 1kg",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "STAR Shea & Turmeric oil 60ml",
        "qty": 5,
        "amount": 110000
      },
      {
        "name": "STAR Shea & Turmeric oil 100ml",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "STAR Shea & Turmeric oil 250ml",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "STAR Elevate Pure Shea Butter",
        "qty": 4,
        "amount": 88000
      },
      {
        "name": "STAR Elevate Refined Shea Butter",
        "qty": 13,
        "amount": 91000
      },
      {
        "name": "STAR Shea Turmeric Soap",
        "qty": 5,
        "amount": 70000
      },
      {
        "name": "STAR Shea Charcoal Soap",
        "qty": 1,
        "amount": 14000
      },
      {
        "name": "STAR Shea Body Smoothie",
        "qty": 4,
        "amount": 180000
      },
      {
        "name": "Art Collabo Angel On Earth",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Amadilu",
        "qty": 2,
        "amount": 90000
      },
      {
        "name": "Basic Shoulder bag",
        "qty": 2,
        "amount": 60000
      },
      {
        "name": "Book bag V2 Zip",
        "qty": 1,
        "amount": 30000
      },
      {
        "name": "Basic Pouch",
        "qty": 19,
        "amount": 85000
      },
      {
        "name": "Book bag V2",
        "qty": 36,
        "amount": 900000
      },
      {
        "name": "Bottle Sticker",
        "qty": 23,
        "amount": 57500
      },
      {
        "name": "Bridget Coin Pouch",
        "qty": 1,
        "amount": 5000
      },
      {
        "name": "Casual bag",
        "qty": 18,
        "amount": 1440000
      },
      {
        "name": "Casual extra bag (ODM only)",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Classic Wallet",
        "qty": 10,
        "amount": 250000
      },
      {
        "name": "Cup holder",
        "qty": 10,
        "amount": 100000
      },
      {
        "name": "Dolphin (L)",
        "qty": 4,
        "amount": 80000
      },
      {
        "name": "Dolphin (S)",
        "qty": 8,
        "amount": 120000
      },
      {
        "name": "Elephant bag (L)",
        "qty": 5,
        "amount": 600000
      },
      {
        "name": "Elephant bag (S)",
        "qty": 2,
        "amount": 200000
      },
      {
        "name": "EYDU Textbook",
        "qty": 3,
        "amount": 85000
      },
      {
        "name": "Flower keyring",
        "qty": 111,
        "amount": 777000
      },
      {
        "name": "Fortune Fish (L)",
        "qty": 45,
        "amount": 675000
      },
      {
        "name": "Fortune Fish (S)",
        "qty": 64,
        "amount": 640000
      },
      {
        "name": "Fruit Coaster",
        "qty": 30,
        "amount": 300000
      },
      {
        "name": "Gloria bag (L)",
        "qty": 5,
        "amount": 750000
      },
      {
        "name": "Gloria bag (S)",
        "qty": 2,
        "amount": 200000
      },
      {
        "name": "Hair Band",
        "qty": 20,
        "amount": 60000
      },
      {
        "name": "Ribbon Hair Band (L)",
        "qty": 2,
        "amount": 30000
      },
      {
        "name": "Ribbon Hair Band (S)",
        "qty": 1,
        "amount": 7000
      },
      {
        "name": "Head band",
        "qty": 16,
        "amount": 208000
      },
      {
        "name": "Jerry Calender",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Jerry Handkerchief",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Jerry Rain Cover",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Jerry Scarf",
        "qty": 4,
        "amount": 12000
      },
      {
        "name": "Jerry Bottle",
        "qty": 2,
        "amount": 40000
      },
      {
        "name": "Jerrycan Water Bottle",
        "qty": 28,
        "amount": 420000
      },
      {
        "name": "Jerrycan bag",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Key holder",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "K. Lucky bag kitenge",
        "qty": 2,
        "amount": 120000
      },
      {
        "name": "K. Lucky bag plain",
        "qty": 3,
        "amount": 240000
      },
      {
        "name": "Leaf keyring",
        "qty": 13,
        "amount": 65000
      },
      {
        "name": "Make Up Pouch",
        "qty": 12,
        "amount": 220000
      },
      {
        "name": "Man In Black Waist Bag",
        "qty": 2,
        "amount": 360000
      },
      {
        "name": "Men's Wallet",
        "qty": 1,
        "amount": 25000
      },
      {
        "name": "Mini Jerrycan",
        "qty": 46,
        "amount": 138000
      },
      {
        "name": "Nature Bag",
        "qty": 2,
        "amount": 160000
      },
      {
        "name": "Original bag (L)",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Original bag (S)",
        "qty": 2,
        "amount": 340000
      },
      {
        "name": "Pajama",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Paper bag L",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Paper bag M",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Paper bag S",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Pencil Case V1",
        "qty": 26,
        "amount": 260000
      },
      {
        "name": "Picnic bag",
        "qty": 24,
        "amount": 720000
      },
      {
        "name": "Plastic Fish",
        "qty": 23,
        "amount": 69000
      },
      {
        "name": "Pot Gloves",
        "qty": 104,
        "amount": 1040000
      },
      {
        "name": "Present pouch (L)",
        "qty": 24,
        "amount": 360000
      },
      {
        "name": "Present pouch (S)",
        "qty": 28,
        "amount": 280000
      },
      {
        "name": "Phone Case Card",
        "qty": 40,
        "amount": 120000
      },
      {
        "name": "Puff Hair Band (S)",
        "qty": 21,
        "amount": 105000
      },
      {
        "name": "Puff Hair Band (L)",
        "qty": 5,
        "amount": 40000
      },
      {
        "name": "Reflector",
        "qty": 13,
        "amount": 156000
      },
      {
        "name": "Kitenge Reflector",
        "qty": 89,
        "amount": 1335000
      },
      {
        "name": "Rope Head Band",
        "qty": 14,
        "amount": 350000
      },
      {
        "name": "Seahorse (L)",
        "qty": 13,
        "amount": 195000
      },
      {
        "name": "Seahorse (S)",
        "qty": 34,
        "amount": 340000
      },
      {
        "name": "Shoppers bag",
        "qty": 1,
        "amount": 15000
      },
      {
        "name": "Shoppers bag V2",
        "qty": 1,
        "amount": 10000
      },
      {
        "name": "Shoppers bag V3",
        "qty": 20,
        "amount": 300000
      },
      {
        "name": "Mula Pouch",
        "qty": 96,
        "amount": 1152000
      },
      {
        "name": "Stand Pouch",
        "qty": 37,
        "amount": 188000
      },
      {
        "name": "Standard bag",
        "qty": 4,
        "amount": 320000
      },
      {
        "name": "Summer Vendor Mrk",
        "qty": 23,
        "amount": 1150000
      },
      {
        "name": "Summer Hat",
        "qty": 4,
        "amount": 60000
      },
      {
        "name": "SDC Square Coaster",
        "qty": 7,
        "amount": 42000
      },
      {
        "name": "SDC Apron",
        "qty": 14,
        "amount": 490000
      },
      {
        "name": "Fruit Coaster Poate",
        "qty": 4,
        "amount": 40000
      },
      {
        "name": "Table Mats Poate",
        "qty": 10,
        "amount": 300000
      },
      {
        "name": "Tyde's 1",
        "qty": 1,
        "amount": 95000
      },
      {
        "name": "Tyde's 2",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Tyde's 3",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Table mat",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Tote bag",
        "qty": 1,
        "amount": 60000
      },
      {
        "name": "Travel Tag",
        "qty": 2,
        "amount": 26000
      },
      {
        "name": "Two Pockets Pouch",
        "qty": 8,
        "amount": 200000
      },
      {
        "name": "Turtle (S)",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Turtle (L)",
        "qty": 25,
        "amount": 750000
      },
      {
        "name": "Turtle XL",
        "qty": 4,
        "amount": 400000
      },
      {
        "name": "T- shirt with one print extra",
        "qty": 2,
        "amount": 76000
      },
      {
        "name": "T-shirt with more prints",
        "qty": 2,
        "amount": 82000
      },
      {
        "name": "T- shirt 2 (SDC)",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "T-shirt",
        "qty": 51,
        "amount": 955000
      },
      {
        "name": "Yakana Creatives 1",
        "qty": 7,
        "amount": 420000
      },
      {
        "name": "Yakana Creatives 2 (S)",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Yakana Creatives 2 (L)",
        "qty": 2,
        "amount": 80000
      },
      {
        "name": "Yakana Creatives 3",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Upcycled Tote bag",
        "qty": 4,
        "amount": 320000
      },
      {
        "name": "Upcycled Shoulder bag",
        "qty": 2,
        "amount": 300000
      },
      {
        "name": "Upcycled Chess bag",
        "qty": 2,
        "amount": 300000
      },
      {
        "name": "Dried Fruits (B)",
        "qty": 9,
        "amount": 91000
      },
      {
        "name": "Dried Fruits (S)",
        "qty": 12,
        "amount": 70000
      },
      {
        "name": "Xmas Stockings",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Xmas deco Bell",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Xmas deco Star",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Xmas deco Tree",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Xmas deco Giftbox",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Xmas deco Ball",
        "qty": 0,
        "amount": 0
      },
      {
        "name": "Whale (L)",
        "qty": 15,
        "amount": 300000
      },
      {
        "name": "Whale (S)",
        "qty": 28,
        "amount": 420000
      },
      {
        "name": "Wrist Bag (S)",
        "qty": 48,
        "amount": 720000
      },
      {
        "name": "Wrist Bag (L)",
        "qty": 35,
        "amount": 875000
      },
      {
        "name": "Wrist pouch",
        "qty": 14,
        "amount": 280000
      },
      {
        "name": "Fruit Coaster (Pumpkin)",
        "qty": 3,
        "amount": 30000
      },
      {
        "name": "Fruit Coaster (Watermelon)",
        "qty": 28,
        "amount": 280000
      }
    ]
  },
  "currentYear": 2026,
  "priorYears": {
    "2025": {
      "channels": {
        "pos": [
          2319199.999928,
          1570000,
          1731600,
          2423000.0999999996,
          4634000,
          3964400,
          2268000,
          6475200,
          2228000.0000004,
          4859600.027000001,
          3988500.050999999,
          8414300.26
        ],
        "custom": [
          0,
          0,
          0,
          2600000,
          0,
          1000000,
          0,
          0,
          0,
          7500000,
          1510000,
          992000
        ],
        "apoyo": [
          300000,
          41000,
          148000,
          710000,
          2310000,
          925000,
          577000,
          2053000,
          453000,
          1355000,
          1695000,
          1752000
        ],
        "bold": [
          351000,
          214500,
          364000,
          0,
          117000,
          208000,
          357500,
          130000,
          0,
          65000,
          22750,
          139750
        ],
        "artisan": [
          0,
          0,
          0,
          0,
          277966,
          344915,
          519492,
          275423.73,
          314407,
          683060,
          830508,
          338983
        ],
        "yujo": [
          0,
          0,
          0,
          0,
          0,
          0,
          0,
          0,
          0,
          0,
          0,
          351400
        ]
      },
      "monthlyGoalUSD": 1000,
      "topProducts": [
        {
          "name": "Elephant bag (L)",
          "qty": 113,
          "amount": 13560000
        },
        {
          "name": "Apoyo Animal (S) + Keyring",
          "qty": 297,
          "amount": 4345000
        },
        {
          "name": "Pajama",
          "qty": 134,
          "amount": 3752000
        },
        {
          "name": "Shoppers bag V2",
          "qty": 100,
          "amount": 3500000
        },
        {
          "name": "Book bag V2",
          "qty": 96,
          "amount": 2446000
        },
        {
          "name": "Casual bag",
          "qty": 26,
          "amount": 2080000
        },
        {
          "name": "Elephant bag (S)",
          "qty": 17,
          "amount": 1660000
        },
        {
          "name": "Standard bag",
          "qty": 19,
          "amount": 1520000
        },
        {
          "name": "Apoyo Eco Bag",
          "qty": 34,
          "amount": 1428000
        },
        {
          "name": "Upcycled Tote bag",
          "qty": 15,
          "amount": 1400000
        }
      ],
      "productPerformance": {
        "totalCount": 137,
        "zeroOrOneCount": 28,
        "lessThan5Count": 54,
        "underperformers": [
          "Apoyo Animal (L)(0)",
          "4Africa New Elephant(0)",
          "4Africa Crocodile(0)",
          "4Africa Rattle(0)",
          "4Africa Boho with straps(0)",
          "4Africa Bracelet(0)",
          "4Africa Dolly without Flap(0)",
          "Coaster(0)",
          "Eg Bag(0)",
          "Key holder(0)",
          "Man In Black Quilted Waist White Bag(0)",
          "Original bag (L)(0)",
          "Paper bag S(0)",
          "Shoppers bag V3(0)",
          "Turtle (S)(0)"
        ]
      },
      "allProducts": [
        {
          "name": "Apoyo Animal (L)",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Apoyo Animal (S)",
          "qty": [
            0,
            2,
            0,
            5,
            15,
            19,
            0,
            6,
            1,
            1,
            4,
            6
          ],
          "amount": [
            0,
            16000,
            0,
            65000,
            195000,
            247000,
            0,
            78000,
            13000,
            13000,
            52000,
            78000
          ]
        },
        {
          "name": "Apoyo Animal (S) + Keyring",
          "qty": [
            20,
            1,
            5,
            14,
            149,
            21,
            4,
            37,
            20,
            11,
            9,
            6
          ],
          "amount": [
            200000,
            10000,
            75000,
            210000,
            2235000,
            315000,
            60000,
            555000,
            300000,
            160000,
            135000,
            90000
          ]
        },
        {
          "name": "Apoyo Hair Band (S)",
          "qty": [
            0,
            0,
            3,
            34,
            9,
            12,
            1,
            34,
            5,
            10,
            9,
            7
          ],
          "amount": [
            0,
            0,
            9000,
            102000,
            27000,
            36000,
            3000,
            102000,
            15000,
            30000,
            27000,
            21000
          ]
        },
        {
          "name": "Apoyo Hair Band (L)",
          "qty": [
            0,
            0,
            0,
            1,
            6,
            3,
            1,
            6,
            0,
            0,
            3,
            5
          ],
          "amount": [
            0,
            0,
            0,
            5000,
            30000,
            15000,
            5000,
            30000,
            0,
            0,
            15000,
            25000
          ]
        },
        {
          "name": "Apoyo Cup Coaster",
          "qty": [
            0,
            0,
            4,
            2,
            0,
            21,
            23,
            6,
            2,
            16,
            14,
            31
          ],
          "amount": [
            0,
            0,
            24000,
            12000,
            0,
            126000,
            138000,
            36000,
            12000,
            96000,
            84000,
            186000
          ]
        },
        {
          "name": "Apoyo Pouch",
          "qty": [
            0,
            0,
            2,
            10,
            0,
            3,
            0,
            16,
            2,
            1,
            0,
            20
          ],
          "amount": [
            0,
            0,
            40000,
            200000,
            0,
            60000,
            0,
            320000,
            40000,
            15000,
            0,
            300000
          ]
        },
        {
          "name": "Apoyo Coin Pouch",
          "qty": [
            0,
            0,
            0,
            1,
            7,
            0,
            1,
            10,
            1,
            6,
            5,
            1
          ],
          "amount": [
            0,
            0,
            0,
            12000,
            84000,
            0,
            12000,
            120000,
            12000,
            72000,
            60000,
            12000
          ]
        },
        {
          "name": "Apoyo Card Pouch",
          "qty": [
            0,
            0,
            0,
            0,
            12,
            0,
            5,
            7,
            0,
            1,
            2,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            252000,
            0,
            105000,
            147000,
            0,
            21000,
            40000,
            0
          ]
        },
        {
          "name": "Apoyo Pencil Case",
          "qty": [
            0,
            0,
            0,
            0,
            2,
            0,
            0,
            3,
            0,
            3,
            0,
            15
          ],
          "amount": [
            0,
            0,
            0,
            0,
            30000,
            0,
            0,
            45000,
            0,
            45000,
            0,
            225000
          ]
        },
        {
          "name": "Apoyo Eco Bag",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            2,
            3,
            4,
            0,
            10,
            4,
            11
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            84000,
            126000,
            168000,
            0,
            420000,
            168000,
            462000
          ]
        },
        {
          "name": "Apoyo Reversable Eco Bag",
          "qty": [
            0,
            0,
            0,
            0,
            4,
            1,
            1,
            0,
            0,
            1,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            168000,
            42000,
            42000,
            0,
            0,
            42000,
            0,
            0
          ]
        },
        {
          "name": "Apoyo Laptop Pouch (L)",
          "qty": [
            0,
            0,
            0,
            2,
            0,
            0,
            0,
            2,
            0,
            1,
            0,
            4
          ],
          "amount": [
            0,
            0,
            0,
            104000,
            0,
            0,
            0,
            104000,
            0,
            52000,
            0,
            208000
          ]
        },
        {
          "name": "Apoyo Laptop Pouch (M)",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0,
            1,
            1,
            2,
            1
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            46000,
            0,
            46000,
            46000,
            92000,
            46000
          ]
        },
        {
          "name": "Apoyo Laptop Pouch (S)",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            2,
            0,
            2,
            1,
            1
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            40000,
            80000,
            0,
            80000,
            40000,
            40000
          ]
        },
        {
          "name": "Apoyo Sports Bag (L)",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            2,
            2,
            29,
            1
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            31000,
            62000,
            62000,
            899000,
            31000
          ]
        },
        {
          "name": "Apoyo Sports Bag (M)",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            4,
            1,
            7,
            2,
            1
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            112000,
            28000,
            196000,
            56000,
            28000
          ]
        },
        {
          "name": "Apoyo Sports Bag (S)",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            5,
            0,
            0,
            1,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            125000,
            0,
            0,
            25000,
            0
          ]
        },
        {
          "name": "4Africa Big Elephant",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            2,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            140000,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "4Africa New Elephant",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "4Africa Big Hippo",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            2,
            0,
            0,
            0,
            2
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            140000,
            0,
            0,
            0,
            140000
          ]
        },
        {
          "name": "4Africa New Monkey",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            2,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            140000,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "4Africa Big Monkey",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            80000,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "4Africa Crocodile",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "4Africa Snake",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            80000,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "4Africa Rattle",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "4Africa Maxi Bunny",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            3,
            0,
            0,
            0,
            1
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            150000,
            0,
            0,
            0,
            50000
          ]
        },
        {
          "name": "4Africa Small Rabbit",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0,
            0,
            1,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            50000,
            0,
            0,
            50000,
            0
          ]
        },
        {
          "name": "4Africa Boho with straps",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "4Africa Kids Apron",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            3,
            2,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            75000,
            50000,
            0,
            0,
            0
          ]
        },
        {
          "name": "4Africa Lion Crochet",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0,
            1
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            70000,
            0,
            70000
          ]
        },
        {
          "name": "4Africa Zebra",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0,
            4
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            70000,
            0,
            280000
          ]
        },
        {
          "name": "4Africa Rhino",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            70000,
            0,
            0
          ]
        },
        {
          "name": "4Africa Gorilla",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            2,
            2
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            30000,
            30000
          ]
        },
        {
          "name": "4Africa Bracelet",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "4Africa Macreme Bag",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            120000,
            0
          ]
        },
        {
          "name": "4Africa Dolly without Flap",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Art Collabo Angel On Earth",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0,
            1
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            150000,
            0,
            150000
          ]
        },
        {
          "name": "Basic Pouch",
          "qty": [
            0,
            0,
            2,
            1,
            1,
            3,
            0,
            3,
            0,
            15,
            20,
            7
          ],
          "amount": [
            0,
            0,
            50000,
            25000,
            25000,
            75000,
            0,
            75000,
            0,
            45000,
            60000,
            21000
          ]
        },
        {
          "name": "Book bag",
          "qty": [
            0,
            0,
            2,
            0,
            0,
            0,
            1,
            0,
            0,
            10,
            8,
            5
          ],
          "amount": [
            0,
            0,
            50000,
            0,
            0,
            0,
            25000,
            0,
            0,
            50000,
            40000,
            25000
          ]
        },
        {
          "name": "Book bag V2",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            2,
            4,
            2,
            9,
            0,
            70,
            9
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            50000,
            130000,
            60000,
            225000,
            0,
            1750000,
            231000
          ]
        },
        {
          "name": "Bottle Bag",
          "qty": [
            1,
            0,
            0,
            0,
            0,
            0,
            3,
            0,
            0,
            2,
            3,
            3
          ],
          "amount": [
            25000,
            0,
            0,
            0,
            0,
            0,
            75000,
            0,
            0,
            20000,
            30000,
            30000
          ]
        },
        {
          "name": "Bottle Sticker",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            2
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            5000
          ]
        },
        {
          "name": "Bottle Ring",
          "qty": [
            0,
            0,
            2,
            0,
            1,
            0,
            0,
            0,
            0,
            0,
            0,
            1
          ],
          "amount": [
            0,
            0,
            20000,
            0,
            10000,
            0,
            0,
            0,
            0,
            0,
            0,
            10000
          ]
        },
        {
          "name": "Bow Tie",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            8,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            40000,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Bridget Coin Pouch",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            5000,
            0
          ]
        },
        {
          "name": "Bridget Hair Clip",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            5000,
            0
          ]
        },
        {
          "name": "Casual bag",
          "qty": [
            2,
            2,
            4,
            3,
            1,
            1,
            2,
            1,
            1,
            3,
            2,
            4
          ],
          "amount": [
            160000,
            160000,
            320000,
            240000,
            80000,
            80000,
            160000,
            80000,
            80000,
            240000,
            160000,
            320000
          ]
        },
        {
          "name": "Casual extra bag",
          "qty": [
            1,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            150000,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Coaster",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Cup holder",
          "qty": [
            2,
            7,
            1,
            2,
            7,
            30,
            0,
            2,
            0,
            9,
            0,
            4
          ],
          "amount": [
            20000,
            70000,
            10000,
            20000,
            70000,
            300000,
            0,
            20000,
            0,
            90000,
            0,
            40000
          ]
        },
        {
          "name": "Cushion",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            3,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            105000,
            0
          ]
        },
        {
          "name": "Canva",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            12
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            360000
          ]
        },
        {
          "name": "Design Tag",
          "qty": [
            3,
            0,
            6,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1
          ],
          "amount": [
            15000,
            0,
            30000,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            3000
          ]
        },
        {
          "name": "Dolphin (L)",
          "qty": [
            0,
            5,
            2,
            5,
            0,
            3,
            1,
            6,
            0,
            0,
            4,
            4
          ],
          "amount": [
            0,
            150000,
            60000,
            150000,
            0,
            80000,
            25000,
            150000,
            0,
            0,
            80000,
            80000
          ]
        },
        {
          "name": "Dolphin (S)",
          "qty": [
            0,
            7,
            0,
            4,
            3,
            1,
            0,
            1,
            0,
            0,
            1,
            2
          ],
          "amount": [
            0,
            140000,
            0,
            80000,
            60000,
            20000,
            0,
            20000,
            0,
            0,
            15000,
            30000
          ]
        },
        {
          "name": "Eg Bag",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Elephant bag (L)",
          "qty": [
            2,
            2,
            2,
            1,
            2,
            0,
            0,
            0,
            1,
            101,
            0,
            2
          ],
          "amount": [
            240000,
            240000,
            240000,
            120000,
            240000,
            0,
            0,
            0,
            120000,
            12120000,
            0,
            240000
          ]
        },
        {
          "name": "Elephant bag (S)",
          "qty": [
            3,
            0,
            0,
            1,
            3,
            2,
            0,
            2,
            1,
            1,
            3,
            1
          ],
          "amount": [
            300000,
            0,
            0,
            100000,
            300000,
            200000,
            0,
            200000,
            100000,
            60000,
            300000,
            100000
          ]
        },
        {
          "name": "EYDU Textbook",
          "qty": [
            0,
            0,
            0,
            1,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0
          ],
          "amount": [
            0,
            0,
            0,
            25000,
            0,
            0,
            0,
            0,
            0,
            0,
            25000,
            0
          ]
        },
        {
          "name": "Flower keyring",
          "qty": [
            0,
            0,
            8,
            14,
            39,
            9,
            2,
            12,
            3,
            5,
            15,
            6
          ],
          "amount": [
            0,
            0,
            56000,
            98000,
            273000,
            63000,
            14000,
            84000,
            21000,
            35000,
            105000,
            42000
          ]
        },
        {
          "name": "Fortune Fish (L)",
          "qty": [
            1,
            0,
            1,
            9,
            6,
            6,
            2,
            6,
            1,
            0,
            1,
            9
          ],
          "amount": [
            15000,
            0,
            15000,
            135000,
            90000,
            90000,
            30000,
            85000,
            15000,
            0,
            15000,
            135000
          ]
        },
        {
          "name": "Fortune Fish (S)",
          "qty": [
            3,
            8,
            1,
            10,
            3,
            0,
            3,
            4,
            1,
            1,
            6,
            13
          ],
          "amount": [
            30000,
            80000,
            10000,
            100000,
            30000,
            0,
            30000,
            40000,
            10000,
            10000,
            60000,
            130000
          ]
        },
        {
          "name": "Fruit Coaster",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            4,
            25
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            40000,
            250000
          ]
        },
        {
          "name": "Gloria bag (L)",
          "qty": [
            0,
            0,
            0,
            0,
            4,
            0,
            0,
            0,
            0,
            1,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            600000,
            0,
            0,
            0,
            0,
            150000,
            0,
            0
          ]
        },
        {
          "name": "Gloria bag (S)",
          "qty": [
            1,
            0,
            1,
            0,
            3,
            1,
            1,
            2,
            0,
            1,
            1,
            1
          ],
          "amount": [
            100000,
            0,
            100000,
            0,
            300000,
            100000,
            100000,
            200000,
            0,
            100000,
            100000,
            100000
          ]
        },
        {
          "name": "Hair Band",
          "qty": [
            4,
            2,
            4,
            1,
            8,
            6,
            3,
            8,
            1,
            5,
            5,
            31
          ],
          "amount": [
            20000,
            10000,
            20000,
            5000,
            40000,
            30000,
            15000,
            40000,
            5000,
            15000,
            15000,
            93000
          ]
        },
        {
          "name": "Hair Band V2",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            24,
            3,
            5,
            1,
            21
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            144000,
            18000,
            26000,
            5000,
            105000
          ]
        },
        {
          "name": "Ribbon Hair Band (L)",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            5,
            1,
            1,
            4,
            1
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            75000,
            15000,
            15000,
            60000,
            15000
          ]
        },
        {
          "name": "Ribbon Hair Band (S)",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            3
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            21000
          ]
        },
        {
          "name": "Head band",
          "qty": [
            3,
            3,
            1,
            0,
            6,
            1,
            2,
            4,
            0,
            0,
            1,
            10
          ],
          "amount": [
            39000,
            39000,
            13000,
            0,
            78000,
            13000,
            26000,
            52000,
            0,
            0,
            13000,
            130000
          ]
        },
        {
          "name": "Jerry Calender",
          "qty": [
            1,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            35000,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Jerry Handkerchief",
          "qty": [
            1,
            0,
            0,
            0,
            2,
            0,
            0,
            1,
            0,
            2,
            0,
            4
          ],
          "amount": [
            3000,
            0,
            0,
            0,
            6000,
            0,
            0,
            3000,
            0,
            6000,
            0,
            12000
          ]
        },
        {
          "name": "Jerry Rain Cover",
          "qty": [
            0,
            1,
            0,
            0,
            1,
            0,
            0,
            1,
            1,
            0,
            0,
            0
          ],
          "amount": [
            0,
            15000,
            0,
            0,
            15000,
            0,
            0,
            15000,
            15000,
            0,
            0,
            0
          ]
        },
        {
          "name": "Jerry Scaf",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            3
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            9000
          ]
        },
        {
          "name": "Jerry Bottle",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            2
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            40000
          ]
        },
        {
          "name": "Jerrycan Water Bottle",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            1,
            0,
            0,
            0,
            0,
            2,
            18
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            10000,
            0,
            0,
            0,
            0,
            30000,
            270000
          ]
        },
        {
          "name": "Jerrycan bag",
          "qty": [
            1,
            0,
            0,
            0,
            1,
            0,
            0,
            0,
            1,
            1,
            0,
            1
          ],
          "amount": [
            80000,
            0,
            0,
            0,
            80000,
            0,
            0,
            0,
            80000,
            80000,
            0,
            80000
          ]
        },
        {
          "name": "Key holder",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "K. Lucky bag kitenge",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0,
            1,
            2,
            1
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            80000,
            0,
            60000,
            120000,
            60000
          ]
        },
        {
          "name": "K. Lucky bag plain",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            2,
            1,
            1,
            1,
            3,
            3,
            2
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            180000,
            100000,
            100000,
            100000,
            260000,
            240000,
            160000
          ]
        },
        {
          "name": "Kitenge Reflector",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            17,
            2,
            8,
            1,
            10
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            25000,
            255000,
            30000,
            120000,
            15000,
            150000
          ]
        },
        {
          "name": "Leaf keyring",
          "qty": [
            0,
            3,
            2,
            4,
            11,
            9,
            0,
            18,
            1,
            3,
            5,
            4
          ],
          "amount": [
            0,
            15000,
            10000,
            20000,
            55000,
            45000,
            0,
            90000,
            5000,
            15000,
            25000,
            20000
          ]
        },
        {
          "name": "Logo",
          "qty": [
            0,
            0,
            0,
            1,
            0,
            0,
            0,
            0,
            0,
            0,
            90,
            0
          ],
          "amount": [
            0,
            0,
            0,
            500000,
            0,
            0,
            0,
            0,
            0,
            0,
            270000,
            0
          ]
        },
        {
          "name": "Man In Black Quilted Waist White Bag",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Mini Jerrycan",
          "qty": [
            0,
            0,
            0,
            17,
            4,
            6,
            8,
            12,
            3,
            17,
            11,
            34
          ],
          "amount": [
            0,
            0,
            0,
            51000,
            12000,
            18000,
            24000,
            36000,
            9000,
            51000,
            33000,
            102000
          ]
        },
        {
          "name": "Nature Bag",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            2
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            80000,
            160000
          ]
        },
        {
          "name": "Original bag (L)",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Original bag (S)",
          "qty": [
            1,
            1,
            1,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0,
            0
          ],
          "amount": [
            120000,
            120000,
            170000,
            0,
            0,
            0,
            0,
            0,
            0,
            170000,
            0,
            0
          ]
        },
        {
          "name": "Pajama",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            134
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            3752000
          ]
        },
        {
          "name": "Paper bag L",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            10,
            10
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            60000,
            60000
          ]
        },
        {
          "name": "Paper bag M",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            50,
            0,
            0,
            0,
            0,
            17
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            200000,
            0,
            0,
            0,
            0,
            68000
          ]
        },
        {
          "name": "Paper bag S",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Pencil Case V1",
          "qty": [
            1,
            1,
            1,
            2,
            0,
            0,
            4,
            5,
            0,
            1,
            6,
            6
          ],
          "amount": [
            20000,
            20000,
            20000,
            40000,
            0,
            0,
            80000,
            100000,
            0,
            20000,
            60000,
            60000
          ]
        },
        {
          "name": "Pencil Case V2",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            4,
            0,
            1,
            1,
            4
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            80000,
            0,
            20000,
            20000,
            80000
          ]
        },
        {
          "name": "Picnic bag",
          "qty": [
            2,
            0,
            1,
            1,
            4,
            0,
            1,
            2,
            1,
            2,
            2,
            17
          ],
          "amount": [
            80000,
            0,
            40000,
            40000,
            160000,
            0,
            40000,
            80000,
            40000,
            55000,
            60000,
            510000
          ]
        },
        {
          "name": "Plastic Fish",
          "qty": [
            13,
            0,
            6,
            0,
            4,
            3,
            1,
            30,
            2,
            1,
            2,
            16
          ],
          "amount": [
            39000,
            0,
            18000,
            0,
            12000,
            9000,
            3000,
            90000,
            6000,
            3000,
            6000,
            48000
          ]
        },
        {
          "name": "Pot Gloves",
          "qty": [
            12,
            2,
            1,
            22,
            0,
            8,
            4,
            12,
            2,
            15,
            19,
            24
          ],
          "amount": [
            120000,
            20000,
            10000,
            220000,
            0,
            80000,
            40000,
            120000,
            20000,
            95000,
            150000,
            240000
          ]
        },
        {
          "name": "Pot Mat (L)",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            13000,
            0,
            0,
            0
          ]
        },
        {
          "name": "Pot Mat (S)",
          "qty": [
            3,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            30000,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Present pouch (L)",
          "qty": [
            1,
            0,
            0,
            0,
            3,
            1,
            2,
            1,
            0,
            0,
            0,
            2
          ],
          "amount": [
            25000,
            0,
            0,
            0,
            75000,
            25000,
            50000,
            25000,
            0,
            0,
            0,
            30000
          ]
        },
        {
          "name": "Present pouch (S)",
          "qty": [
            6,
            4,
            0,
            0,
            2,
            6,
            1,
            2,
            0,
            3,
            9,
            21
          ],
          "amount": [
            120000,
            80000,
            0,
            0,
            40000,
            120000,
            20000,
            40000,
            0,
            60000,
            90000,
            210000
          ]
        },
        {
          "name": "Phone Case Card",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            5,
            1,
            5,
            3,
            7
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            35000,
            7000,
            12000,
            9000,
            29000
          ]
        },
        {
          "name": "Reflector",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            2,
            4,
            1,
            1,
            1,
            1,
            6
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            24000,
            48000,
            12000,
            12000,
            12000,
            12000,
            75000
          ]
        },
        {
          "name": "Rope Head Band",
          "qty": [
            2,
            1,
            2,
            1,
            7,
            2,
            0,
            1,
            3,
            6,
            1,
            3
          ],
          "amount": [
            50000,
            25000,
            50000,
            25000,
            175000,
            50000,
            0,
            25000,
            75000,
            150000,
            25000,
            75000
          ]
        },
        {
          "name": "Sample_Bag",
          "qty": [
            0,
            0,
            0,
            1,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            200000,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Sample_Flower & leaf keyring",
          "qty": [
            7,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            35000,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Sample_pouch",
          "qty": [
            1,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            15000,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Seahorse (L)",
          "qty": [
            0,
            3,
            1,
            1,
            3,
            3,
            0,
            0,
            1,
            0,
            5,
            8
          ],
          "amount": [
            0,
            75000,
            25000,
            25000,
            75000,
            75000,
            0,
            0,
            25000,
            0,
            75000,
            120000
          ]
        },
        {
          "name": "Seahorse (S)",
          "qty": [
            0,
            1,
            0,
            1,
            0,
            0,
            0,
            2,
            0,
            1,
            6,
            6
          ],
          "amount": [
            0,
            15000,
            0,
            15000,
            0,
            0,
            0,
            30000,
            0,
            15000,
            60000,
            60000
          ]
        },
        {
          "name": "Shoppers bag",
          "qty": [
            2,
            1,
            1,
            0,
            2,
            0,
            3,
            0,
            0,
            1,
            1,
            10
          ],
          "amount": [
            70000,
            35000,
            35000,
            0,
            70000,
            0,
            105000,
            0,
            0,
            15000,
            15000,
            150000
          ]
        },
        {
          "name": "Shoppers bag V2",
          "qty": [
            0,
            0,
            0,
            100,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            3500000,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Shoppers bag V3",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Square Pouch",
          "qty": [
            2,
            1,
            1,
            0,
            0,
            0,
            0,
            0,
            0,
            2,
            0,
            0
          ],
          "amount": [
            40000,
            20000,
            20000,
            0,
            0,
            0,
            0,
            0,
            0,
            40000,
            0,
            0
          ]
        },
        {
          "name": "Stand Pouch",
          "qty": [
            0,
            0,
            5,
            4,
            0,
            1,
            0,
            2,
            0,
            13,
            23,
            2
          ],
          "amount": [
            0,
            0,
            120000,
            100000,
            0,
            25000,
            0,
            50000,
            0,
            39000,
            69000,
            6000
          ]
        },
        {
          "name": "Standard bag",
          "qty": [
            0,
            1,
            0,
            0,
            0,
            0,
            0,
            0,
            2,
            0,
            2,
            14
          ],
          "amount": [
            0,
            80000,
            0,
            0,
            0,
            0,
            0,
            0,
            160000,
            0,
            160000,
            1120000
          ]
        },
        {
          "name": "Summer Hat",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            3
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            45000
          ]
        },
        {
          "name": "Tyde's",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0,
            0,
            1
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            65000,
            0,
            0,
            95000
          ]
        },
        {
          "name": "Tyde's 2",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            100000,
            0
          ]
        },
        {
          "name": "Table mat",
          "qty": [
            1,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1
          ],
          "amount": [
            25000,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            15000
          ]
        },
        {
          "name": "Tote bag",
          "qty": [
            1,
            0,
            0,
            2,
            2,
            1,
            0,
            1,
            0,
            0,
            0,
            4
          ],
          "amount": [
            80000,
            0,
            0,
            160000,
            160000,
            80000,
            0,
            80000,
            0,
            0,
            0,
            240000
          ]
        },
        {
          "name": "Travel Tag",
          "qty": [
            2,
            0,
            0,
            1,
            16,
            0,
            1,
            2,
            1,
            2,
            0,
            3
          ],
          "amount": [
            30000,
            0,
            0,
            15000,
            240000,
            0,
            13000,
            26000,
            13000,
            26000,
            0,
            39000
          ]
        },
        {
          "name": "Two Pockets Pouch",
          "qty": [
            5,
            0,
            0,
            2,
            0,
            0,
            1,
            0,
            0,
            0,
            0,
            6
          ],
          "amount": [
            140000,
            0,
            0,
            56000,
            0,
            0,
            28000,
            0,
            0,
            0,
            0,
            150000
          ]
        },
        {
          "name": "Turtle (S)",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Turtle (L)",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            9,
            4,
            5,
            8,
            10
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            270000,
            160000,
            150000,
            240000,
            300000
          ]
        },
        {
          "name": "Turtle XL",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            0,
            1,
            0,
            4
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            80000,
            0,
            100000,
            0,
            400000
          ]
        },
        {
          "name": "Upcycled Tote bag",
          "qty": [
            5,
            0,
            0,
            0,
            3,
            2,
            0,
            1,
            0,
            1,
            0,
            3
          ],
          "amount": [
            500000,
            0,
            0,
            0,
            300000,
            200000,
            0,
            100000,
            0,
            60000,
            0,
            240000
          ]
        },
        {
          "name": "Upcycled Shoulder bag",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            2,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            300000,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Upcycled Chess bag",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            2,
            1,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            300000,
            150000,
            0
          ]
        },
        {
          "name": "Dried Fruits (B)",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            4,
            0,
            18,
            8,
            1
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            60000,
            0,
            234000,
            104000,
            13000
          ]
        },
        {
          "name": "Dried Fruits (S)",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            6,
            9,
            9,
            4,
            1
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            72000,
            92000,
            90000,
            40000,
            10000
          ]
        },
        {
          "name": "Xmas Stockings",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Xmas deco Bell",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            3
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            30000
          ]
        },
        {
          "name": "Xmas deco Star",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            8000
          ]
        },
        {
          "name": "Xmas deco Tree",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            3
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            24000
          ]
        },
        {
          "name": "Xmas deco Giftbox",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            7
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            37000
          ]
        },
        {
          "name": "Xmas deco Ball",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        },
        {
          "name": "Whale (L)",
          "qty": [
            0,
            3,
            2,
            8,
            2,
            0,
            1,
            3,
            0,
            0,
            2,
            4
          ],
          "amount": [
            0,
            75000,
            50000,
            200000,
            50000,
            0,
            25000,
            75000,
            0,
            0,
            40000,
            80000
          ]
        },
        {
          "name": "Whale (S)",
          "qty": [
            0,
            3,
            0,
            6,
            3,
            2,
            9,
            3,
            1,
            0,
            2,
            13
          ],
          "amount": [
            0,
            60000,
            0,
            120000,
            60000,
            40000,
            180000,
            60000,
            20000,
            0,
            30000,
            195000
          ]
        },
        {
          "name": "Wrist bag",
          "qty": [
            2,
            0,
            1,
            4,
            2,
            1,
            2,
            2,
            5,
            2,
            1,
            10
          ],
          "amount": [
            74000,
            0,
            37000,
            148000,
            74000,
            37000,
            74000,
            74000,
            185000,
            74000,
            25000,
            250000
          ]
        },
        {
          "name": "Wrist pouch",
          "qty": [
            0,
            0,
            4,
            7,
            2,
            0,
            0,
            3,
            0,
            1,
            0,
            0
          ],
          "amount": [
            0,
            0,
            100000,
            175000,
            50000,
            0,
            0,
            75000,
            0,
            20000,
            0,
            0
          ]
        },
        {
          "name": "Workshop",
          "qty": [
            0,
            0,
            0,
            0,
            0,
            1,
            0,
            0,
            0,
            0,
            0,
            0
          ],
          "amount": [
            0,
            0,
            0,
            0,
            0,
            1000000,
            0,
            0,
            0,
            0,
            0,
            0
          ]
        }
      ],
      "paymentStats": [
        {
          "method": "Cash",
          "amount": 26588200.407928
        },
        {
          "method": "Bank",
          "amount": 19847000
        },
        {
          "method": "Card",
          "amount": 10508200
        },
        {
          "method": "MTN Mobile Money",
          "amount": 4592600
        },
        {
          "method": "Airtel Mobile Money",
          "amount": 1860800.03
        }
      ],
      "posInvoiceCounts": [
        27,
        15,
        20,
        38,
        35,
        34,
        26,
        51,
        27,
        52,
        48,
        143
      ],
      "visits": [
        36,
        85,
        57,
        86,
        100,
        99,
        77,
        93,
        116,
        111,
        89,
        136
      ],
      "contacts": [
        17,
        32,
        27,
        9,
        27,
        17,
        27,
        39,
        44,
        59,
        43,
        61
      ],
      "sold": [
        27,
        53,
        81,
        124,
        193,
        188,
        105,
        395,
        96,
        203,
        291,
        555
      ],
      "dowStats": [
        {
          "day": "Mon",
          "visits": 153,
          "contacts": 56,
          "revenue": 5069200
        },
        {
          "day": "Tue",
          "visits": 142,
          "contacts": 48,
          "revenue": 3720200
        },
        {
          "day": "Wed",
          "visits": 142,
          "contacts": 50,
          "revenue": 4799000
        },
        {
          "day": "Thu",
          "visits": 129,
          "contacts": 63,
          "revenue": 3162300
        },
        {
          "day": "Fri",
          "visits": 191,
          "contacts": 66,
          "revenue": 11125200
        },
        {
          "day": "Sat",
          "visits": 205,
          "contacts": 84,
          "revenue": 7020900
        },
        {
          "day": "Sun",
          "visits": 123,
          "contacts": 35,
          "revenue": 6403200
        }
      ],
      "visitorOrigin": [
        {
          "origin": "Korea",
          "count": 105
        },
        {
          "origin": "Uganda",
          "count": 87
        },
        {
          "origin": "Netherlands",
          "count": 10
        },
        {
          "origin": "Sudan",
          "count": 10
        },
        {
          "origin": "USA",
          "count": 9
        },
        {
          "origin": "Germany",
          "count": 8
        },
        {
          "origin": "India",
          "count": 8
        },
        {
          "origin": "China",
          "count": 6
        },
        {
          "origin": "Other",
          "count": 32
        }
      ],
      "visitorType": [
        {
          "type": "Returning / Loyal Customers",
          "count": 49,
          "color": "#2F6E68"
        },
        {
          "type": "New / First-time Visitors",
          "count": 3,
          "color": "#D69A2D"
        }
      ],
      "eventKeywords": [
        {
          "keyword": "market",
          "label": "Market/Event",
          "count": 31
        },
        {
          "keyword": "training",
          "label": "Training Program",
          "count": 48
        },
        {
          "keyword": "church",
          "label": "Church",
          "count": 11
        },
        {
          "keyword": "koica",
          "label": "KOICA",
          "count": 4
        },
        {
          "keyword": "recommend(ed)? by|referr?(ed|al)|friend of",
          "label": "Referral",
          "count": 14
        }
      ],
      "operationStats": {
        "operatingDays": 341,
        "zeroVisitDays": 60,
        "zeroVisitPct": 0.17595307917888564,
        "zeroSalesDays": 140,
        "zeroSalesPct": 0.41055718475073316,
        "top5Concentration": 0.1565181598062954,
        "totalRevenue": 41300000
      }
    }
  },
  "lastUploadedAt": "2026-09-29T08:22:16.561Z"
};

const PAY_COLORS = [COLORS.teal, COLORS.ochre, COLORS.clay, COLORS.slate, COLORS.olive];
const VENDOR_COLORS = [COLORS.teal, COLORS.ochre, COLORS.clay, COLORS.slate, COLORS.olive, COLORS.plum, "#8AA0A8"];

const STORAGE_KEY = "sdc_giftshop_dashboard_v1";

/* ---------------------------------- helpers ---------------------------------- */
const fmtUGX = (n) => `${Math.round(n).toLocaleString("en-US")} UGX`;
const fmtCompact = (n) => {
  if (Math.abs(n) >= 1000000) return `${(n / 1000000).toFixed(1)}M`;
  if (Math.abs(n) >= 1000) return `${(n / 1000).toFixed(0)}K`;
  return `${Math.round(n)}`;
};
const fmtNum = (n) => Math.round(n).toLocaleString("en-US");
const fmtPct = (n) => `${(n * 100).toFixed(1)}%`;

function monthTotal(data, i) {
  return REVENUE_KEYS.reduce((s, k) => s + (Number(data.channels[k][i]) || 0), 0);
}
function activeMonthCount(data) {
  let n = 0;
  for (let i = 0; i < 12; i++) if (monthTotal(data, i) > 0) n = i + 1;
  return n;
}
function activeMonthsForArray(arr) {
  let n = 0;
  for (let i = 0; i < arr.length; i++) if (Number(arr[i] || 0) > 0) n = i + 1;
  return n;
}

function linreg(values) {
  const n = values.length;
  if (n < 2) return { slope: 0, intercept: values[0] || 0 };
  const sumX = (n * (n - 1)) / 2;
  const sumY = values.reduce((a, b) => a + b, 0);
  const sumXY = values.reduce((a, y, i) => a + i * y, 0);
  const sumXX = values.reduce((a, _, i) => a + i * i, 0);
  const denom = n * sumXX - sumX * sumX;
  const slope = denom === 0 ? 0 : (n * sumXY - sumX * sumY) / denom;
  const intercept = (sumY - slope * sumX) / n;
  return { slope, intercept };
}

function computeForecast(monthlyTotals, activeMonths, annualGoalUGX) {
  if (activeMonths === 0) return null;
  const actualSeries = monthlyTotals.slice(0, activeMonths);
  const { slope, intercept } = linreg(actualSeries);
  const ytd = actualSeries.reduce((a, b) => a + b, 0);
  const avgMonthly = ytd / activeMonths;
  const remaining = 12 - activeMonths;
  const projectedFlat = ytd + avgMonthly * remaining;
  const trendSeries = MONTH_LABELS.map((label, i) => {
    const row = { period: label };
    if (i < activeMonths) row.actual = monthlyTotals[i];
    else row.trend = Math.max(0, slope * i + intercept);
    if (i === activeMonths - 1) row.trend = monthlyTotals[i];
    return row;
  });
  const requiredAvgRemaining = remaining > 0 ? Math.max(0, (annualGoalUGX - ytd) / remaining) : 0;
  const lastMonthRev = monthlyTotals[activeMonths - 1];
  const prevMonthRev = activeMonths >= 2 ? monthlyTotals[activeMonths - 2] : null;
  const momGrowth = prevMonthRev ? (lastMonthRev - prevMonthRev) / prevMonthRev : null;
  return {
    ytd, avgMonthly, remaining, projectedFlat, trendSeries, annualGoalUGX,
    requiredAvgRemaining, momGrowth, projectedAchievement: annualGoalUGX > 0 ? projectedFlat / annualGoalUGX : 0,
  };
}

function buildHighlights(data) {
  const origin = data.visitorOrigin || [];
  const type = data.visitorType || [];
  const events = [...(data.eventKeywords || [])].sort((a, b) => b.count - a.count);
  const lines = [];
  if (origin.length >= 2) {
    lines.push(`Visitors most mentioned in notes: ${origin[0].origin} (${origin[0].count} mentions), followed by ${origin[1].origin} (${origin[1].count} mentions).`);
  }
  if (type.length >= 2 && type[1].count > 0) {
    const ratio = (type[0].count / type[1].count).toFixed(1);
    lines.push(`Mentions of returning/loyal customers are about ${ratio}x mentions of new/first-time visitors — a strong base of repeat customers.`);
  }
  if (events.length) {
    const top = events.slice(0, 3).map((e) => `${e.label} (${e.count})`).join(", ");
    lines.push(`Most frequent keywords in notes: ${top}`);
  }
  return lines;
}

/* ---------------------------------- excel parsing engine ---------------------------------- */
// 매월 원본 엑셀 파일(일일 리포트 / 판매 데이터)을 그대로 다시 업로드하면
// 아래 함수들이 각 시트를 읽어 대시보드 데이터 구조로 변환합니다.

const MONTH_CODES = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
const DOW_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
const DOW_KR = { Monday: "Mon", Tuesday: "Tue", Wednesday: "Wed", Thursday: "Thu", Friday: "Fri", Saturday: "Sat", Sunday: "Sun" };
const NATIONALITY_MAP = [
  { key: "ugandan", label: "Uganda" }, { key: "korean", label: "Korea" }, { key: "german", label: "Germany" },
  { key: "american", label: "USA" }, { key: "chinese", label: "China" }, { key: "italian", label: "Italy" },
  { key: "french", label: "France" }, { key: "japanese", label: "Japan" }, { key: "nigerian", label: "Nigeria" },
  { key: "kenyan", label: "Kenya" }, { key: "dutch", label: "Netherlands" }, { key: "eritrean", label: "Eritrea" },
  { key: "ethiopian", label: "Ethiopia" }, { key: "sudanese", label: "Sudan" }, { key: "british", label: "UK" },
  { key: "canadian", label: "Canada" }, { key: "indian", label: "India" },
];
const PAYMENT_LABELS = { cash: "Cash", card: "Card", mtn: "MTN Mobile Money", airtel: "Airtel Mobile Money", bank: "Bank" };

function countMatches(text, pattern) {
  const m = text.match(new RegExp(pattern, "gi"));
  return m ? m.length : 0;
}

// 연도별 파일마다 'Date' 헤더 셀이 오타·누락되는 경우가 있어, 헤더 이름 대신
// 실제 값이 대부분 날짜(Date)인 컬럼을 내용 기반으로 찾아내는 범용 리더
// 아티팩트가 iframe 안에서 실행될 때, 부모 창/iframe이 서로 다른 realm의 Date 클래스를 가질 수 있어
// 'instanceof Date'가 실패할 수 있음 (특히 라이브러리가 다른 realm에서 Date 객체를 만든 경우).
// Object.prototype.toString 기반 체크는 realm에 상관없이 안전하게 동작함.
function isDateLike(v) {
  return v != null && Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime ? v.getTime() : NaN);
}

// 엑셀 날짜 일련번호 → Date 변환 과정에서 부동소수점 오차로 자정이 전날 23:59:59 등으로 미세하게
// 밀리는 경우가 있어(환경에 따라 다르게 나타남), 가장 가까운 UTC 자정으로 반올림해서 날짜 경계 오류를 방지함
function normalizeDateToUTCMidnight(d) {
  if (!isDateLike(d)) return d;
  const rounded = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  const hours = d.getUTCHours() + d.getUTCMinutes() / 60 + d.getUTCSeconds() / 3600 + d.getUTCMilliseconds() / 3600000;
  if (hours >= 12) rounded.setUTCDate(rounded.getUTCDate() + 1);
  return rounded;
}

function buildRowReader(ws, knownHeaders) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  let headerRowIdx = -1, bestMatches = -1, colIndex = {};
  for (let r = 0; r < Math.min(rows.length, 10); r++) {
    const row = rows[r] || [];
    const idx = {};
    let matches = 0;
    row.forEach((cell, ci) => {
      if (cell == null) return;
      const s = String(cell).trim().toLowerCase();
      knownHeaders.forEach((h) => { if (s === h.toLowerCase() && idx[h] === undefined) { idx[h] = ci; matches++; } });
    });
    if (matches > bestMatches) { bestMatches = matches; headerRowIdx = r; colIndex = idx; }
  }
  if (headerRowIdx < 0 || bestMatches <= 0) return null;
  let dateCol = -1, bestScore = 0;
  const scanEnd = Math.min(rows.length, headerRowIdx + 60);
  const numCols = rows.slice(headerRowIdx + 1, scanEnd).reduce((m, r) => Math.max(m, r ? r.length : 0), 0);
  for (let c = 0; c < numCols; c++) {
    let dateCount = 0, total = 0;
    for (let r = headerRowIdx + 1; r < scanEnd; r++) {
      const v = rows[r] && rows[r][c];
      if (v == null) continue;
      total++;
      if (isDateLike(v)) dateCount++;
    }
    if (total > 3 && dateCount / total > bestScore) { bestScore = dateCount / total; dateCol = c; }
  }
  if (bestScore > 0.5) colIndex.__date__ = dateCol;
  return { rows, headerRowIdx, colIndex };
}

function findMonthHeaderRow(rows) {
  for (let r = 0; r < Math.min(rows.length, 15); r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      if (row[c] != null && String(row[c]).trim().toUpperCase() === "JAN") return { r, c };
    }
  }
  return null;
}

function parseMonthlySalesSheet(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const hit = findMonthHeaderRow(rows);
  if (!hit) throw new Error("Could not find month (JAN~DEC) headers. Please check the '2026 Monthly Sales' sheet format.");
  const headerRow = rows[hit.r];
  const monthCols = {};
  let searchFrom = hit.c;
  MONTH_CODES.forEach((m) => {
    const idx = headerRow.findIndex((cell, ci) => ci >= searchFrom && cell != null && String(cell).trim().toUpperCase() === m);
    if (idx >= 0) { monthCols[m] = idx; searchFrom = idx + 1; }
  });
  const labelCol = Math.max(0, hit.c - 1);

  function findRow(test) {
    for (let r = hit.r + 1; r < rows.length; r++) {
      const cell = rows[r] ? rows[r][labelCol] : null;
      if (cell != null && test(String(cell).toLowerCase())) return rows[r];
    }
    return null;
  }
  function extractSeries(row) {
    if (!row) return MONTH_CODES.map(() => 0);
    return MONTH_CODES.map((m) => {
      const idx = monthCols[m];
      const v = idx != null ? row[idx] : null;
      return typeof v === "number" ? v : 0;
    });
  }

  const channels = {
    pos: extractSeries(findRow((l) => l.includes("pos"))),
    custom: extractSeries(findRow((l) => l.includes("custom order"))),
    apoyo: extractSeries(findRow((l) => l.includes("apoyo"))),
    bold: extractSeries(findRow((l) => l.includes("bold"))),
    artisan: extractSeries(findRow((l) => l.includes("artisan"))),
    yujo: extractSeries(findRow((l) => l.includes("yujo"))),
  };

  // SDC TOTAL 행의 수식(예: =SUM(C5,C8:C11))을 읽어, 이미 아는 채널(POS/BOLD/Artisan/Yujo) 외에
  // 새로 추가된 외부 입점업체 행(예: SAWE)이 있으면 자동으로 찾아서 giftshop 채널로 추가함.
  // 수식에 없는 행(Custom Order/Apoyo처럼 셀에 색칠된 행)은 POS에 이미 포함된 위탁판매로 보고 건너뜀.
  // 주의: sheet_to_json의 배열 인덱스는 시트의 실제 사용 범위(!ref)가 A열부터 시작하지 않으면
  // 실제 엑셀 열 문자와 어긋나므로, 반드시 range 시작 오프셋을 더해서 셀 주소를 계산해야 함.
  const newChannels = [];
  let totalRowIdx = -1;
  for (let r = hit.r + 1; r < rows.length; r++) {
    const cell = rows[r] ? rows[r][labelCol] : null;
    const label = cell != null ? String(cell) : "";
    if (/total\s*\(ugx\)/i.test(label) || (/\btotal\b/i.test(label) && !/goal|krw|usd|%/i.test(label))) { totalRowIdx = r; break; }
  }
  if (totalRowIdx >= 0 && ws) {
    const range = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : { s: { r: 0, c: 0 } };
    const rowOffset = range.s.r;
    const colOffset = range.s.c;
    const firstMonthKey = MONTH_CODES.find((m) => monthCols[m] != null);
    const firstMonthCol = firstMonthKey != null ? monthCols[firstMonthKey] : null;
    if (firstMonthCol != null) {
      const addr = XLSX.utils.encode_cell({ r: totalRowIdx + rowOffset, c: firstMonthCol + colOffset });
      const totalCell = ws[addr];
      const formula = totalCell && totalCell.f ? totalCell.f : null;
      if (formula) {
        const refs = formula.match(/[A-Z]+\d+(:[A-Z]+\d+)?/g) || [];
        const includedRows = new Set();
        refs.forEach((ref) => {
          if (ref.includes(":")) {
            const [a, b] = ref.split(":");
            const ca = XLSX.utils.decode_cell(a), cb = XLSX.utils.decode_cell(b);
            for (let r = ca.r; r <= cb.r; r++) includedRows.add(r - rowOffset);
          } else {
            includedRows.add(XLSX.utils.decode_cell(ref).r - rowOffset);
          }
        });
        const knownKeywords = ["pos", "custom order", "apoyo", "bold", "artisan", "yujo"];
        for (let r = hit.r + 1; r < totalRowIdx; r++) {
          const cell2 = rows[r] ? rows[r][labelCol] : null;
          const label2 = cell2 != null ? String(cell2).trim() : "";
          if (!label2) continue;
          const lower = label2.toLowerCase();
          if (knownKeywords.some((kw) => lower.includes(kw))) continue; // 이미 아는 채널
          if (!includedRows.has(r)) continue; // 색칠된(제외) 행 — 위탁판매로 보고 건너뜀
          const key = lower.replace(/[^a-z0-9]+/g, "").slice(0, 20) || `ch${r}`;
          if (!channels[key]) {
            channels[key] = extractSeries(rows[r]);
            newChannels.push({ key, label: label2.replace(/\n/g, " ") });
          }
        }
      }
    }
  }

  let monthlyGoalUSD = null;
  const goalRow = findRow((l) => l.includes("goal"));
  if (goalRow) {
    const vals = extractSeries(goalRow).filter((v) => v > 0);
    if (vals.length) monthlyGoalUSD = vals[0];
  }

  // 헤더 행에서 "20XX TOTAL" 같은 문구나, 시트 상단 제목(예: "▶ 2026 MONTHLY SALES REPORT")에서 연도 추출
  let year = null;
  for (const cell of headerRow) {
    if (cell == null) continue;
    const m = String(cell).match(/(20\d{2})/);
    if (m) { year = Number(m[1]); break; }
  }
  if (year == null) {
    for (let r = 0; r < Math.min(rows.length, hit.r); r++) {
      const row = rows[r] || [];
      for (const cell of row) {
        if (cell == null) continue;
        const m = String(cell).match(/(20\d{2})/);
        if (m) { year = Number(m[1]); break; }
      }
      if (year != null) break;
    }
  }

  return { channels, monthlyGoalUSD, monthsFound: Object.keys(monthCols).length, year, newChannels };
}

function parseAllProductsSheet(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const hit = findMonthHeaderRow(rows);
  if (!hit) return [];
  const headerRow = rows[hit.r];
  const monthCols = {};
  let searchFrom = hit.c;
  MONTH_CODES.forEach((m) => {
    const idx = headerRow.findIndex((cell, ci) => ci >= searchFrom && cell != null && String(cell).trim().toUpperCase() === m);
    if (idx >= 0) { monthCols[m] = idx; searchFrom = idx + 1; }
  });
  const labelCol = Math.max(0, hit.c - 1);
  const products = [];
  for (let r = hit.r + 2; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row[labelCol] == null || String(row[labelCol]).trim() === "") continue;
    const label = String(row[labelCol]).trim();
    if (/subtotal|discount|^total$/i.test(label)) break;
    const qty = MONTH_CODES.map((m) => {
      const idx = monthCols[m];
      const v = idx != null ? row[idx] : null;
      return typeof v === "number" ? v : 0;
    });
    const amount = MONTH_CODES.map((m) => {
      const idx = monthCols[m];
      const v = idx != null ? row[idx + 1] : null;
      return typeof v === "number" ? v : 0;
    });
    products.push({ name: label.replace(/\n/g, " "), qty, amount });
  }
  return products;
}

function parseTopProductsSheet(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const hit = findMonthHeaderRow(rows);
  if (!hit) return { topProducts: [], productPerformance: null };
  const headerRow = rows[hit.r];
  let totalQtyCol = -1;
  for (let c = hit.c; c < headerRow.length; c++) {
    if (headerRow[c] != null && /20\d{2}\s*TOTAL/i.test(String(headerRow[c]))) { totalQtyCol = c; break; }
  }
  if (totalQtyCol < 0) return { topProducts: [], productPerformance: null };
  const totalAmountCol = totalQtyCol + 1;
  const labelCol = Math.max(0, hit.c - 1);
  const products = [];
  for (let r = hit.r + 2; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row[labelCol] == null || String(row[labelCol]).trim() === "") continue;
    const label = String(row[labelCol]).trim();
    if (/subtotal|discount|^total$/i.test(label)) break;
    let qty = typeof row[totalQtyCol] === "number" ? row[totalQtyCol] : null;
    let amount = typeof row[totalAmountCol] === "number" ? row[totalAmountCol] : null;
    // 이 행에 TOTAL 칸(수식/값)이 비어있으면 — 월별 QTY/Gross 칸을 직접 더해서 계산함
    // (새로 추가된 상품 행에 TOTAL 수식이 안 끌어와져 있는 경우가 있어, 0으로 잘못 표시되는 걸 방지)
    if (qty == null || amount == null) {
      let sumQty = 0, sumAmount = 0;
      for (let i = 0; i < 12; i++) {
        const qCol = hit.c + i * 2, gCol = qCol + 1;
        if (typeof row[qCol] === "number") sumQty += row[qCol];
        if (typeof row[gCol] === "number") sumAmount += row[gCol];
      }
      if (qty == null) qty = sumQty;
      if (amount == null) amount = sumAmount;
    }
    products.push({ name: label.replace(/\n/g, " "), qty, amount });
  }
  const topProducts = [...products].sort((a, b) => b.amount - a.amount).slice(0, 10);

  // 저조 판매 상품 분석 — 이미 단종/샘플/서비스성 항목은 제외하고 실제 판매 후보만 집계
  const excludePattern = /discontinued|sample_|workshop|repair service|^logo$/i;
  const sellable = products.filter((p) => !excludePattern.test(p.name));
  const zeroOrOne = sellable.filter((p) => p.qty <= 1);
  const lessThan5 = sellable.filter((p) => p.qty < 5);
  const underperformers = [...lessThan5].sort((a, b) => a.qty - b.qty).slice(0, 15).map((p) => `${p.name}(${p.qty})`);
  const productPerformance = {
    totalCount: sellable.length,
    zeroOrOneCount: zeroOrOne.length,
    lessThan5Count: lessThan5.length,
    underperformers,
    allSellable: sellable,
  };
  return { topProducts, productPerformance };
}

function normalizePaymentMethod(raw) {
  const l = raw.toLowerCase();
  if (l.includes("cash")) return "cash";
  if (l.includes("card")) return "card";
  if (l.includes("mtn")) return "mtn";
  if (l.includes("airtel")) return "airtel";
  if (l.includes("bank")) return "bank";
  return raw;
}

// 상품 변형(예: "T-shirt", "T-shirt with one print extra")을 사용자가 지정한 그룹으로 묶어서
// 판매수량/금액을 합쳐줌. products: [{name, qty, amount}], groups: [{canonicalName, members:[...]}]
// 공백/대소문자 차이(예: "T-shirt" vs "T- shirt")에 흔들리지 않도록 정규화한 키로 비교함
function normalizeProductNameKey(name) {
  return String(name || "").replace(/\s+/g, "").toLowerCase();
}

function applyProductGrouping(products, groups) {
  if (!groups || groups.length === 0) return products;
  const memberKeyToGroup = {};
  groups.forEach((g) => { (g.members || []).forEach((m) => { memberKeyToGroup[normalizeProductNameKey(m)] = g.canonicalName; }); });
  const merged = {};
  const order = [];
  products.forEach((p) => {
    const key = memberKeyToGroup[normalizeProductNameKey(p.name)] || p.name;
    if (!merged[key]) { merged[key] = { name: key, qty: 0, amount: 0 }; order.push(key); }
    merged[key].qty += Number(p.qty) || 0;
    merged[key].amount += Number(p.amount) || 0;
  });
  return order.map((k) => merged[k]);
}

function normalizeVendorName(raw) {
  const v = String(raw).trim();
  if (v.toLowerCase() === "star") return "STAR";
  return v;
}

function parsePosDataSheet(ws) {
  const reader = buildRowReader(ws, ["Invoice No.", "Item", "Rate", "Qty", "Gross Total", "Discount Rate", "Discount Amount", "Net Total", "Payment mode", "Category"]);
  if (!reader || reader.colIndex.__date__ == null) return null;
  const { rows, headerRowIdx, colIndex } = reader;
  const dateCol = colIndex.__date__;
  const invoiceCol = colIndex["Invoice No."], itemCol = colIndex["Item"], qtyCol = colIndex["Qty"], grossCol = colIndex["Gross Total"],
    discountCol = colIndex["Discount Amount"], netCol = colIndex["Net Total"], paymentCol = colIndex["Payment mode"],
    categoryCol = colIndex["Category"];

  const records = [];
  for (let r = headerRowIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const d = normalizeDateToUTCMidnight(row[dateCol]);
    if (!isDateLike(d)) continue;
    records.push({
      date: d,
      invoice: invoiceCol != null ? row[invoiceCol] : null,
      item: itemCol != null ? row[itemCol] : null,
      qty: qtyCol != null ? Number(row[qtyCol]) || 0 : 0,
      gross: grossCol != null ? Number(row[grossCol]) || 0 : 0,
      discount: discountCol != null ? Number(row[discountCol]) || 0 : 0,
      net: netCol != null ? Number(row[netCol]) || 0 : 0,
      payment: paymentCol != null ? row[paymentCol] : null,
      category: categoryCol != null ? row[categoryCol] : null,
    });
  }
  if (records.length === 0) return null;
  const yearCounts = {};
  records.forEach((r) => { const y = r.date.getUTCFullYear(); yearCounts[y] = (yearCounts[y] || 0) + 1; });
  const year = Number(Object.entries(yearCounts).sort((a, b) => b[1] - a[1])[0][0]);
  const yearRecords = records.filter((r) => r.date.getUTCFullYear() === year);

  const paymentTotals = {};
  yearRecords.forEach((r) => {
    if (!r.payment) return;
    const norm = normalizePaymentMethod(String(r.payment).trim());
    paymentTotals[norm] = (paymentTotals[norm] || 0) + r.net;
  });
  const paymentStats = Object.entries(paymentTotals)
    .map(([m, amount]) => ({ method: PAYMENT_LABELS[m] || m, amount }))
    .sort((a, b) => b.amount - a.amount);

  const qty = Array(12).fill(0), net = Array(12).fill(0);
  const invoiceSets = Array.from({ length: 12 }, () => new Set());
  yearRecords.forEach((r) => {
    const m = r.date.getUTCMonth();
    qty[m] += r.qty; net[m] += r.net;
    if (r.invoice) invoiceSets[m].add(r.invoice);
  });
  const posInvoiceCounts = invoiceSets.map((s) => s.size);

  let vendorResult = null;
  if (categoryCol != null) {
    const byVendor = {}, monthlyByVendor = {}, monthlyDetail = {};
    yearRecords.forEach((r) => {
      if (!r.category) return;
      const v = normalizeVendorName(r.category);
      if (!byVendor[v]) byVendor[v] = { vendor: v, qty: 0, gross: 0, net: 0, discount: 0, count: 0 };
      byVendor[v].qty += r.qty; byVendor[v].gross += r.gross; byVendor[v].net += r.net; byVendor[v].discount += r.discount; byVendor[v].count += 1;
      const m = r.date.getUTCMonth();
      if (!monthlyByVendor[v]) monthlyByVendor[v] = Array(12).fill(0);
      monthlyByVendor[v][m] += r.net;
      if (!monthlyDetail[v]) monthlyDetail[v] = Array.from({ length: 12 }, () => ({ qty: 0, gross: 0, net: 0, discount: 0, count: 0 }));
      monthlyDetail[v][m].qty += r.qty; monthlyDetail[v][m].gross += r.gross; monthlyDetail[v][m].net += r.net; monthlyDetail[v][m].discount += r.discount; monthlyDetail[v][m].count += 1;
    });
    vendorResult = { vendorStats: Object.values(byVendor).sort((a, b) => b.net - a.net), monthlyByVendor, monthlyDetail };
  }

  // 상품명 -> 생산업체(카테고리) 매핑 (저조 판매 상품을 업체별로 나눠 보기 위함)
  let itemVendorMap = null;
  if (categoryCol != null && itemCol != null) {
    const itemVendorCounts = {};
    yearRecords.forEach((r) => {
      if (!r.item || !r.category) return;
      const key = normalizeProductNameKey(r.item);
      const v = normalizeVendorName(r.category);
      if (!itemVendorCounts[key]) itemVendorCounts[key] = {};
      itemVendorCounts[key][v] = (itemVendorCounts[key][v] || 0) + 1;
    });
    itemVendorMap = {};
    Object.entries(itemVendorCounts).forEach(([key, counts]) => {
      itemVendorMap[key] = Object.entries(counts).sort((a, b) => b[1] - a[1])[0][0];
    });
  }

  // 상품별 월별 수량·순매출(Net) 집계 — Details Report는 할인 전(Gross) 금액만 있어 할인이
  // 상품마다 다르게 적용되면 실제 순매출과 어긋날 수 있어, POS Data의 거래별 Net Total로 다시 집계함
  let productResult = null;
  if (itemCol != null) {
    const byProduct = {};
    yearRecords.forEach((r) => {
      if (!r.item) return;
      const name = String(r.item).trim();
      if (!name) return;
      if (!byProduct[name]) byProduct[name] = { name, qty: Array(12).fill(0), amount: Array(12).fill(0) };
      const m = r.date.getUTCMonth();
      byProduct[name].qty[m] += r.qty;
      byProduct[name].amount[m] += r.net;
    });
    const allProducts = Object.values(byProduct);
    const cumulative = allProducts.map((p) => ({
      name: p.name, qty: p.qty.reduce((a, b) => a + b, 0), amount: p.amount.reduce((a, b) => a + b, 0),
    }));
    const excludePattern = /discontinued|sample_|workshop|repair service|^logo$/i;
    const sellable = cumulative.filter((p) => !excludePattern.test(p.name));
    const zeroOrOne = sellable.filter((p) => p.qty <= 1);
    const lessThan5 = sellable.filter((p) => p.qty < 5);
    const underperformers = [...lessThan5].sort((a, b) => a.qty - b.qty).slice(0, 15).map((p) => `${p.name}(${p.qty})`);
    const topProducts = [...cumulative].sort((a, b) => b.amount - a.amount).slice(0, 10);
    productResult = {
      allProducts, topProducts,
      productPerformance: { totalCount: sellable.length, zeroOrOneCount: zeroOrOne.length, lessThan5Count: lessThan5.length, underperformers, allSellable: sellable },
    };
  }

  return { year, paymentStats, posMonthlyQty: qty, posMonthlyNet: net, posInvoiceCounts, vendorResult, itemVendorMap, productResult, yearRecords };
}

function parseVisitorText(allNotes) {
  const originCounts = NATIONALITY_MAP
    .map((n) => ({ origin: n.label, count: countMatches(allNotes, "\\b" + n.key + "\\w*") }))
    .filter((o) => o.count > 0)
    .sort((a, b) => b.count - a.count);
  const top = originCounts.slice(0, 8);
  const restSum = originCounts.slice(8).reduce((a, o) => a + o.count, 0);
  const visitorOrigin = restSum > 0 ? [...top, { origin: "Other", count: restSum }] : top;

  const returning = countMatches(allNotes, "old customer") + countMatches(allNotes, "re-?visit")
    + countMatches(allNotes, "regular( visitor| customer)?") + countMatches(allNotes, "return customer")
    + countMatches(allNotes, "returning");
  const newV = countMatches(allNotes, "first time") + countMatches(allNotes, "never visited")
    + countMatches(allNotes, "new customer") + countMatches(allNotes, "first visit");
  const visitorType = [
    { type: "Returning / Loyal Customers", count: returning, color: "#2F6E68" },
    { type: "New / First-time Visitors", count: newV, color: "#D69A2D" },
  ];

  const eventDefs = [
    { keyword: "market", label: "Market/Event" },
    { keyword: "training", label: "Training Program" },
    { keyword: "church", label: "Church" },
    { keyword: "koica", label: "KOICA" },
    { keyword: "recommend(ed)? by|referr?(ed|al)|friend of", label: "Referral" },
  ];
  const eventKeywords = eventDefs.map((e) => ({ keyword: e.keyword, label: e.label, count: countMatches(allNotes, e.keyword) }));

  return { visitorOrigin, visitorType, eventKeywords };
}

// 일부 파일에 이미 월별로 집계된 피봇 테이블 시트가 있음 (예: "SUM of Total Visits" 헤더).
// 거래/일별 로우를 하나하나 날짜로 재분류하는 방식이 일부 환경에서 불안정한 것이 확인되어,
// 이런 피봇 요약이 있으면 그걸 우선 신뢰해서 방문자/판매수량을 계산함.
const MONTH_ABBR_MAP = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, sept: 8, oct: 9, nov: 10, dec: 11 };
function findVisitsPivotSheet(wb) {
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true });
    for (let r = 0; r < Math.min(rows.length, 5); r++) {
      const row = rows[r] || [];
      const visitsCol = row.findIndex((c) => typeof c === "string" && /sum of total visits/i.test(c));
      if (visitsCol < 0) continue;
      const revenueCol = row.findIndex((c) => typeof c === "string" && /sum of sales revenue/i.test(c));
      const soldCol = row.findIndex((c) => typeof c === "string" && /sum of total sold/i.test(c));
      const visits = Array(12).fill(0), revenue = Array(12).fill(0), sold = Array(12).fill(0);
      let found = false;
      for (let rr = r + 1; rr < rows.length; rr++) {
        const label = rows[rr] && rows[rr][0] != null ? String(rows[rr][0]).trim().toLowerCase() : "";
        const m = MONTH_ABBR_MAP[label];
        if (m == null) continue;
        found = true;
        if (typeof rows[rr][visitsCol] === "number") visits[m] = rows[rr][visitsCol];
        if (revenueCol >= 0 && typeof rows[rr][revenueCol] === "number") revenue[m] = rows[rr][revenueCol];
        if (soldCol >= 0 && typeof rows[rr][soldCol] === "number") sold[m] = rows[rr][soldCol];
      }
      if (found) return { visits, revenue, sold };
    }
  }
  return null;
}

function parseDailyReportSheet(ws) {
  const reader = buildRowReader(ws, ["Day", "Total Visits", "Total Contacts", "Total Sold", "Sales Revenue", "Notes"]);
  if (!reader || reader.colIndex.__date__ == null) {
    return { visits: Array(12).fill(0), contacts: Array(12).fill(0), sold: Array(12).fill(0), dowStats: [], visitorOrigin: [], visitorType: [], eventKeywords: [], monthsWithData: 0, operationStats: null, year: null };
  }
  const { rows, headerRowIdx, colIndex } = reader;
  const dateCol = colIndex.__date__;
  const dayCol = colIndex["Day"], visitsCol = colIndex["Total Visits"], contactsCol = colIndex["Total Contacts"],
    soldCol = colIndex["Total Sold"], revenueCol = colIndex["Sales Revenue"], notesCol = colIndex["Notes"];

  const dated = [];
  for (let r = headerRowIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const d = normalizeDateToUTCMidnight(row[dateCol]);
    if (!isDateLike(d)) continue;
    dated.push({
      date: d,
      day: dayCol != null ? row[dayCol] : null,
      visits: visitsCol != null ? Number(row[visitsCol]) || 0 : 0,
      contacts: contactsCol != null ? Number(row[contactsCol]) || 0 : 0,
      sold: soldCol != null ? Number(row[soldCol]) || 0 : 0,
      revenue: revenueCol != null ? Number(row[revenueCol]) || 0 : 0,
      notes: notesCol != null ? row[notesCol] : null,
    });
  }
  if (dated.length === 0) {
    return { visits: Array(12).fill(0), contacts: Array(12).fill(0), sold: Array(12).fill(0), dowStats: [], visitorOrigin: [], visitorType: [], eventKeywords: [], monthsWithData: 0, operationStats: null, year: null };
  }
  const yearCounts = {};
  dated.forEach((d) => { const y = d.date.getUTCFullYear(); yearCounts[y] = (yearCounts[y] || 0) + 1; });
  const year = Number(Object.entries(yearCounts).sort((a, b) => b[1] - a[1])[0][0]);

  const monthly = Array.from({ length: 12 }, () => ({ visits: 0, contacts: 0, sold: 0, revenue: 0 }));
  const dow = {};
  DOW_ORDER.forEach((d) => (dow[d] = { visits: 0, contacts: 0, revenue: 0 }));
  const notesArr = [];
  let monthsWithData = 0;
  dated.forEach((d) => {
    if (d.date.getUTCFullYear() !== year) return;
    const m = d.date.getUTCMonth();
    monthly[m].visits += d.visits;
    monthly[m].contacts += d.contacts;
    monthly[m].sold += d.sold;
    monthly[m].revenue += d.revenue;
    if (d.day && dow[d.day]) {
      dow[d.day].visits += d.visits;
      dow[d.day].contacts += d.contacts;
      dow[d.day].revenue += d.revenue;
    }
    if (d.notes) notesArr.push(String(d.notes));
  });
  monthly.forEach((m) => { if (m.visits > 0 || m.revenue > 0) monthsWithData += 1; });

  const yearDated = dated.filter((d) => d.date.getUTCFullYear() === year);
  const activeDays = yearDated.filter((d) => d.visits > 0 || d.revenue > 0);
  let operationStats = null;
  if (activeDays.length > 0) {
    const maxTime = Math.max(...activeDays.map((d) => d.date.getTime()));
    const inRange = yearDated.filter((d) => d.date.getTime() <= maxTime);
    const zeroVisitDays = inRange.filter((d) => d.visits === 0).length;
    const zeroSalesDays = inRange.filter((d) => d.revenue === 0).length;
    const revenues = inRange.map((d) => d.revenue).sort((a, b) => b - a);
    const totalRev = revenues.reduce((a, b) => a + b, 0);
    const top5Concentration = totalRev > 0 ? revenues.slice(0, 5).reduce((a, b) => a + b, 0) / totalRev : 0;
    operationStats = {
      operatingDays: inRange.length,
      zeroVisitDays, zeroVisitPct: inRange.length > 0 ? zeroVisitDays / inRange.length : 0,
      zeroSalesDays, zeroSalesPct: inRange.length > 0 ? zeroSalesDays / inRange.length : 0,
      top5Concentration, totalRevenue: totalRev,
    };
  }

  const visits = monthly.map((m) => m.visits);
  const contacts = monthly.map((m) => m.contacts);
  const sold = monthly.map((m) => m.sold);
  const dowStats = DOW_ORDER.map((d) => ({ day: DOW_KR[d], visits: dow[d].visits, contacts: dow[d].contacts, revenue: dow[d].revenue }));
  const { visitorOrigin, visitorType, eventKeywords } = parseVisitorText(notesArr.join("\n"));
  return { visits, contacts, sold, dowStats, visitorOrigin, visitorType, eventKeywords, monthsWithData, operationStats, year };
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("Could not read the file."));
    reader.readAsArrayBuffer(file);
  });
}

function findSheetName(sheetNames, keyword) {
  const lower = keyword.toLowerCase();
  return sheetNames.find((n) => n.toLowerCase().includes(lower)) || null;
}

// Reads an optional "Custom Orders" tab (same workbook as Monthly Sales/POS Data) so the order
// log no longer has to be hand-typed into the dashboard. Expected columns: Month (1-12, or a
// date column), Customer, Category, Product, Qty, Amount, Note, and an optional Group column for
// tagging multiple product lines that belong to one real order under the same group value.
function parseCustomOrdersSheet(ws) {
  const reader = buildRowReader(ws, ["Month", "Customer", "Category", "Product", "Qty", "Amount", "Note", "Group"]);
  if (!reader) return null;
  const { rows, headerRowIdx, colIndex } = reader;
  const monthCol = colIndex["Month"], customerCol = colIndex["Customer"], categoryCol = colIndex["Category"],
    productCol = colIndex["Product"], qtyCol = colIndex["Qty"], amountCol = colIndex["Amount"],
    noteCol = colIndex["Note"], groupCol = colIndex["Group"];
  const dateCol = colIndex["Date"] != null ? colIndex["Date"] : colIndex.__date__;
  if (customerCol == null) return null;
  const orders = [];
  for (let r = headerRowIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const customer = row[customerCol];
    if (customer == null || String(customer).trim() === "") continue;
    let month = 0;
    if (monthCol != null && typeof row[monthCol] === "number") {
      month = Math.min(11, Math.max(0, Math.round(row[monthCol]) - 1));
    } else if (dateCol != null && isDateLike(row[dateCol])) {
      month = normalizeDateToUTCMidnight(row[dateCol]).getUTCMonth();
    }
    const id = r - headerRowIdx;
    const groupRaw = groupCol != null ? row[groupCol] : null;
    orders.push({
      id,
      groupId: groupRaw != null && String(groupRaw).trim() !== "" ? String(groupRaw).trim() : id,
      month,
      customer: String(customer).trim(),
      category: categoryCol != null && row[categoryCol] != null && String(row[categoryCol]).trim() !== "" ? String(row[categoryCol]).trim() : null,
      product: productCol != null && row[productCol] != null ? String(row[productCol]).trim() : "",
      qty: qtyCol != null && typeof row[qtyCol] === "number" ? row[qtyCol] : null,
      amount: amountCol != null && typeof row[amountCol] === "number" ? row[amountCol] : 0,
      note: noteCol != null && row[noteCol] != null ? String(row[noteCol]) : "",
    });
  }
  return orders.length > 0 ? orders : null;
}

// Reads an optional "Custom Order Leads" (or just "Leads") tab: Customer + Description columns,
// combined into the same "Customer: description" strings the Leads list already displays.
function parseCustomOrderLeadsSheet(ws) {
  const reader = buildRowReader(ws, ["Customer", "Description", "Detail", "Inquiry"]);
  if (!reader) return null;
  const { rows, headerRowIdx, colIndex } = reader;
  const customerCol = colIndex["Customer"];
  if (customerCol == null) return null;
  const descCol = colIndex["Description"] != null ? colIndex["Description"]
    : (colIndex["Detail"] != null ? colIndex["Detail"] : colIndex["Inquiry"]);
  const leads = [];
  for (let r = headerRowIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const customer = row[customerCol];
    if (customer == null || String(customer).trim() === "") continue;
    const desc = descCol != null && row[descCol] != null ? String(row[descCol]).trim() : "";
    leads.push(desc ? `${String(customer).trim()}: ${desc}` : String(customer).trim());
  }
  return leads.length > 0 ? leads : null;
}

// Reads free-text memos ("메모"/Note in Google Sheets, exported into the xlsx as legacy/threaded
// cell comments) attached to each month's cell on the "Custom Order" row of the Monthly Sales
// sheet. This is an alternative to a separate "Custom Orders" tab for staff who'd rather keep
// writing in the same memo box they already use: each order is one line, "Customer / Category /
// Product / Qty / Amount" (any field but Customer can be left blank, e.g. "kilta / Hotel/Restaurant
// / / / 858000"), and a still-open inquiry is written as "LEAD: ...". Multiple lines that repeat
// the same customer name within one month's memo are combined into a single order (their
// products/quantities/amounts are merged) since one real order shouldn't count as several.
// CUSTOMER_CATEGORIES(파일 상단, UI에서 쓰는 카테고리 목록)는 엑셀 파싱 엔진 추출 범위 밖에 있어
// 여기서는 별도로 복사해둠 -- 카테고리 목록을 바꾸면 두 곳 다 맞춰줘야 함.
const MEMO_CATEGORY_NAMES = ["Hotel/Restaurant", "Government Office", "NGO/Institution", "Individual", "Other"];
const MEMO_LEAD_PREFIX_RE = /^lead\s*:\s*/i;
const CATEGORY_SLASH_PLACEHOLDER = "⁄"; // stands in for "/" inside category names like "Hotel/Restaurant" while splitting on "/"

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// xlsx(SheetJS)가 레거시 코멘트(cell.c[].t)를 읽을 때, 원문이 아스키가 아니면(한글 등) UTF-8 바이트를
// 한 바이트=한 글자로 잘못 디코딩하는 버그가 있어("적용" -> "ì ì©") 여기서 직접 복구함. 이미 정상
// 디코딩된 텍스트(전부 아스키이거나, 애초에 깨지지 않은 텍스트)에는 영향이 없음.
function fixMojibakeUtf8(str) {
  if (typeof TextDecoder === "undefined") return str;
  try {
    const bytes = new Uint8Array(str.length);
    for (let i = 0; i < str.length; i++) {
      const code = str.charCodeAt(i);
      if (code > 255) return str;
      bytes[i] = code;
    }
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch (e) {
    return str;
  }
}

function extractMemoText(cell) {
  if (!cell || !cell.c || cell.c.length === 0) return "";
  const raw = cell.c.map((c) => (c && c.t ? c.t : "")).join("\n");
  let text = fixMojibakeUtf8(raw);
  // 구버전 엑셀 호환 안내문이 앞에 붙어있는 경우 제거 (신버전 xlsx 내보내기에서는 대개 이미 없음)
  text = text.replace(/\[Threaded comment\][\s\S]*?Comment:\s*\n?/i, "");
  // 쓰레드 답글(해결됨/다시 열림 등 상태 메모)은 실제 주문 내용이 아니므로 첫 "Reply:" 이전까지만 사용
  const replyIdx = text.search(/\n\s*reply\s*:/i);
  if (replyIdx >= 0) text = text.slice(0, replyIdx);
  return text.trim();
}

function protectCategorySlashes(line) {
  let out = line;
  MEMO_CATEGORY_NAMES.forEach((cat) => {
    if (!cat.includes("/")) return;
    const re = new RegExp(escapeRegExp(cat), "gi");
    out = out.replace(re, cat.replace(/\//g, CATEGORY_SLASH_PLACEHOLDER));
  });
  return out;
}

function parseMemoOrderLine(line) {
  const protectedLine = protectCategorySlashes(line);
  if (!protectedLine.includes("/")) return null;
  const restore = (s) => s.trim().replace(new RegExp(CATEGORY_SLASH_PLACEHOLDER, "g"), "/");
  const parts = protectedLine.split("/").map(restore);
  const customer = parts[0];
  if (!customer) return null;
  // 옛날 자유 텍스트 메모(예: "9/14 72,000 received(cash)")도 "/"가 들어있어 잘못 구조화된 줄로
  // 오인될 수 있음 -- 2번째 칸이 정해진 카테고리 이름 5개 중 하나와 정확히 일치할 때만 새 형식의
  // 줄로 인정해서, 날짜 등 우연히 "/"가 들어간 옛날 메모는 안전하게 걸러냄(= unparsed 처리).
  const categoryRaw = parts[1] || "";
  const category = MEMO_CATEGORY_NAMES.find((cat) => cat.toLowerCase() === categoryRaw.toLowerCase());
  if (!category) return null;
  const product = parts[2] && parts[2] !== "" ? parts[2] : "";
  const qtyRaw = parts[3] != null ? parts[3] : "";
  const amountRaw = parts[4] != null ? parts[4] : "";
  const qtyNum = qtyRaw !== "" ? Number(qtyRaw.replace(/[^0-9.\-]/g, "")) : NaN;
  const amountNum = amountRaw !== "" ? Number(amountRaw.replace(/[^0-9.\-]/g, "")) : NaN;
  return {
    customer,
    category,
    product,
    qty: Number.isFinite(qtyNum) ? qtyNum : null,
    amount: Number.isFinite(amountNum) ? amountNum : 0,
  };
}

function parseCustomOrderMemoSheet(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const hit = findMonthHeaderRow(rows);
  if (!hit) {
    return { orders: [], leads: null, unparsedCount: 0, debug: ["⚠️ 월(JAN~DEC) 헤더 행을 찾지 못했습니다."] };
  }
  const headerRow = rows[hit.r];
  const monthCols = {};
  let searchFrom = hit.c;
  MONTH_CODES.forEach((m) => {
    const idx = headerRow.findIndex((cell, ci) => ci >= searchFrom && cell != null && String(cell).trim().toUpperCase() === m);
    if (idx >= 0) { monthCols[m] = idx; searchFrom = idx + 1; }
  });
  const labelCol = Math.max(0, hit.c - 1);
  let customOrderRowIdx = -1;
  for (let r = hit.r + 1; r < rows.length; r++) {
    const cell = rows[r] ? rows[r][labelCol] : null;
    if (cell != null && String(cell).toLowerCase().includes("custom order")) { customOrderRowIdx = r; break; }
  }
  if (customOrderRowIdx < 0 || !ws["!ref"]) {
    return { orders: [], leads: null, unparsedCount: 0, debug: ["⚠️ 'Custom Order' 라벨이 포함된 행을 찾지 못했습니다 (라벨 열의 문구를 확인해주세요)."] };
  }

  const range = XLSX.utils.decode_range(ws["!ref"]);
  const rowOffset = range.s.r, colOffset = range.s.c;

  const orders = [];
  const leads = [];
  let unparsedCount = 0;
  let orderIdCounter = 1;

  // 리드(진행 중인 문의/앞으로 납품해야 하는 논의 중 주문)는 확정 주문과 달리 "현재 상태"를
  // 나타내는 정보라, 여러 달치를 무조건 계속 누적하면 이미 해결된 옛날 문의까지 계속 남아있게
  // 되어 오히려 부정확해짐. 다만 담당자가 이번 달 미확정 주문뿐 아니라 앞으로 몇 달 뒤 납품
  // 예정인 주문까지 미리 해당 달 메모에 나눠 적어두는 경우가 있어서, "실제로 주문 수치가 찍힌
  // 마지막 달"을 기준으로 그 달과 그 이후 달(아직 오지 않은 미래 달 포함)의 LEAD: 줄은 전부
  // 반영함 -- 시간이 지나 그 미래 달에 실제 주문 수치가 채워지면 기준 달이 자동으로 그만큼
  // 앞으로 이동해서, 이미 지난 달에 남아있던 LEAD: 줄은 자연스럽게 더 이상 반영되지 않음.
  let lastOrderMonth = -1;
  let anyRecognized = false;
  MONTH_CODES.forEach((m, month) => {
    const col = monthCols[m];
    if (col == null) return;
    const addr = XLSX.utils.encode_cell({ r: customOrderRowIdx + rowOffset, c: col + colOffset });
    const text = extractMemoText(ws[addr]);
    if (!text) return;
    text.split("\n").forEach((line) => {
      const trimmed = line.trim();
      if (!trimmed) return;
      if (MEMO_LEAD_PREFIX_RE.test(trimmed)) { anyRecognized = true; return; }
      if (parseMemoOrderLine(trimmed) != null) { anyRecognized = true; lastOrderMonth = month; }
    });
  });
  const leadStartMonth = Math.max(lastOrderMonth, 0);

  // 진단용: 실제로 자동화가 월별 Custom Order 메모(노트)를 어떻게 읽어들였는지 실행 로그에
  // 남기기 위한 정보. "메모 자체가 아예 안 읽히는지"와 "메모는 읽혔지만 형식이 안 맞는지"를
  // 구분할 수 있게 해줌 (구글 시트 → xlsx 내보내기 과정에서 노트가 누락되는 경우가 있어 확인용).
  const monthDebug = [];

  MONTH_CODES.forEach((m, month) => {
    const col = monthCols[m];
    if (col == null) return;
    const addr = XLSX.utils.encode_cell({ r: customOrderRowIdx + rowOffset, c: col + colOffset });
    const text = extractMemoText(ws[addr]);
    if (!text) { monthDebug.push(`${m}:메모없음`); return; }
    const lines = text.split("\n").map((l) => l.trim()).filter((l) => l.length > 0);
    let recognizedThisMonth = 0;
    let leadsThisMonth = 0;
    let unparsedBefore = unparsedCount;
    const groups = new Map(); // 고객명(소문자) -> 주문 하나로 합쳐서 누적
    lines.forEach((line) => {
      if (MEMO_LEAD_PREFIX_RE.test(line)) {
        leadsThisMonth++;
        if (month >= leadStartMonth) leads.push(line.replace(MEMO_LEAD_PREFIX_RE, "").trim());
        return;
      }
      const parsed = parseMemoOrderLine(line);
      if (!parsed) { unparsedCount++; return; }
      recognizedThisMonth++;
      const key = parsed.customer.toLowerCase();
      if (!groups.has(key)) {
        groups.set(key, {
          id: orderIdCounter++,
          groupId: `${month}-${key}`,
          month,
          customer: parsed.customer,
          category: parsed.category,
          product: parsed.product,
          qty: parsed.qty,
          amount: parsed.amount,
          note: "",
        });
      } else {
        const existing = groups.get(key);
        if (parsed.product) existing.product = existing.product ? `${existing.product}, ${parsed.product}` : parsed.product;
        if (parsed.qty != null) existing.qty = (existing.qty || 0) + parsed.qty;
        existing.amount += parsed.amount;
        if (!existing.category && parsed.category) existing.category = parsed.category;
      }
    });
    groups.forEach((o) => orders.push(o));
    monthDebug.push(`${m}:${text.length}자/${lines.length}줄(인식${recognizedThisMonth},리드${leadsThisMonth},미인식${unparsedCount - unparsedBefore})`);
  });

  // anyRecognized가 false면 어떤 달도 아직 새 형식을 안 쓰고 있다는 뜻이라 리드 쪽은 손대지
  // 않음(null). 새 형식을 쓰는 달/LEAD: 줄이 하나라도 있으면, 기준 달 이후에 LEAD: 줄이 하나도
  // 없어도(=문의가 전부 해결됨) 빈 배열을 돌려줘서 예전 리드가 그대로 남지 않고 실제로 비워지게 함.
  return { orders, leads: anyRecognized ? leads : null, unparsedCount, debug: monthDebug };
}

// Reads an optional "Market Events" tab: just Date + Name (or Event) columns -- no revenue
// typing needed. Revenue/qty/invoice counts are computed automatically by summing that date's
// POS Data transactions, replacing the old "color the row orange in POS Data" convention (cell
// background color can't be read reliably from an exported xlsx, but the date+name text can).
function parseMarketEventDatesSheet(ws, yearRecords) {
  const reader = buildRowReader(ws, ["Date", "Name", "Event"]);
  if (!reader) return null;
  const { rows, headerRowIdx, colIndex } = reader;
  // Prefer the literal "Date" header match over the date-value-scanning heuristic in
  // buildRowReader (which needs several rows to trust a column as dates) -- a short list of
  // market days should still work even with just a couple of rows.
  const dateCol = colIndex["Date"] != null ? colIndex["Date"] : colIndex.__date__;
  if (dateCol == null) return null;
  const nameCol = colIndex["Name"] != null ? colIndex["Name"] : colIndex["Event"];
  const byDate = {};
  (yearRecords || []).forEach((rec) => {
    const key = rec.date.toISOString().slice(0, 10);
    if (!byDate[key]) byDate[key] = { revenue: 0, qty: 0, invoices: new Set() };
    byDate[key].revenue += rec.net;
    byDate[key].qty += rec.qty;
    if (rec.invoice) byDate[key].invoices.add(rec.invoice);
  });
  const events = [];
  let id = 1;
  for (let r = headerRowIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const d = row[dateCol];
    if (!isDateLike(d)) continue;
    const dateKey = normalizeDateToUTCMidnight(d).toISOString().slice(0, 10);
    const name = nameCol != null && row[nameCol] != null ? String(row[nameCol]).trim() : "Market Day";
    const agg = byDate[dateKey];
    events.push({
      id: id++,
      date: dateKey,
      month: Number(dateKey.slice(5, 7)) - 1,
      name,
      revenue: agg ? agg.revenue : 0,
      qty: agg ? agg.qty : 0,
      invoices: agg ? agg.invoices.size : 0,
    });
  }
  return events.length > 0 ? events : null;
}

const YEAR_FIELD_KEYS = [
  "channels", "monthlyGoalUSD", "visits", "contacts", "sold", "dowStats", "visitorOrigin", "visitorType",
  "eventKeywords", "operationStats", "topProducts", "productPerformance", "allProducts", "paymentStats",
  "posInvoiceCounts", "posMonthlyQty", "vendorStats", "vendorMonthly", "vendorMonthlyDetail", "itemVendorMap",
];
function extractYearFields(source) {
  const out = {};
  YEAR_FIELD_KEYS.forEach((k) => { if (source[k] !== undefined) out[k] = source[k]; });
  return out;
}

async function processExcelFiles(fileList, referenceYear) {
  const files = Array.from(fileList);
  const bundles = [];
  const messages = [];
  for (const file of files) {
    try {
      const buf = await readFileAsArrayBuffer(file);
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const salesSheetName = findSheetName(wb.SheetNames, "monthly sales");
      const dailySheetName = findSheetName(wb.SheetNames, "daily report");

      if (salesSheetName) {
        const { channels, monthlyGoalUSD, year, newChannels } = parseMonthlySalesSheet(wb.Sheets[salesSheetName]);
        const fields = { channels, monthlyGoalUSD };
        if (newChannels && newChannels.length > 0) {
          fields.newChannels = newChannels;
          messages.push(`✓ New external channel(s) auto-detected: ${newChannels.map((c) => c.label).join(", ")} (included in gift shop revenue)`);
        }
        let activeMonths = 0;
        for (let i = 0; i < 12; i++) {
          const total = Object.keys(channels).reduce((s, k) => s + (Number(channels[k][i]) || 0), 0);
          if (total > 0) activeMonths = i + 1;
        }
        messages.push(`✓ Sales data file recognized${year ? ` (${year}, ` : " ("}${activeMonths} month(s) confirmed)`);

        const detailsSheetName = findSheetName(wb.SheetNames, "details report");
        if (detailsSheetName) {
          const { topProducts, productPerformance } = parseTopProductsSheet(wb.Sheets[detailsSheetName]);
          fields.topProducts = topProducts;
          if (productPerformance) fields.productPerformance = productPerformance;
          fields.allProducts = parseAllProductsSheet(wb.Sheets[detailsSheetName]);
        }
        // 이 파일이 과거(지난) 연도로 확인되면 POS Data 시트는 아예 건드리지 않음 — 지난 연도는
        // Monthly Sales 요약값만 신뢰하고, POS Data는 B2B/B2C가 섞여 있는 등 세부분석에 못 쓸 수 있어
        // 생산업체·결제수단·수량 분석은 진행 연도(currentYear) 파일에만 적용함
        const isHistoricalYear = referenceYear != null && year != null && year < referenceYear;
        const posSheetName = findSheetName(wb.SheetNames, "pos data");
        let posYear = null;
        let posResult = null;
        if (posSheetName && !isHistoricalYear) {
          posResult = parsePosDataSheet(wb.Sheets[posSheetName]);
          if (posResult) {
            posYear = posResult.year;
            fields.paymentStats = posResult.paymentStats;
            fields.posInvoiceCounts = posResult.posInvoiceCounts;
            // 매출(개요 페이지 월별 총계)은 Monthly Sales 시트 요약값을 그대로 신뢰함 (POS Data의
            // 거래별 날짜 재분류 단계가 일부 환경에서 불안정하게 동작하는 것이 확인되어 되돌림).
            // 생산업체(카테고리)·결제 수단 분석은 Monthly Sales로는 볼 수 없어 계속 POS Data를 사용함.
            fields.posMonthlyQty = posResult.posMonthlyQty;
            if (posResult.itemVendorMap) fields.itemVendorMap = posResult.itemVendorMap;
            if (posResult.vendorResult) {
              fields.vendorStats = posResult.vendorResult.vendorStats;
              fields.vendorMonthly = posResult.vendorResult.monthlyByVendor;
              fields.vendorMonthlyDetail = posResult.vendorResult.monthlyDetail;
              messages.push(`✓ Vendor (category) data recognized (${posResult.vendorResult.vendorStats.length} vendors)`);
            }
            // 베스트셀러/저조판매 상품은 Details Report(할인 전 금액)보다 POS Data(할인 반영된 순매출)가
            // 더 정확해서, POS Data 기준으로 다시 계산한 값으로 덮어씀
            if (posResult.productResult) {
              fields.allProducts = posResult.productResult.allProducts;
              fields.topProducts = posResult.productResult.topProducts;
              fields.productPerformance = posResult.productResult.productPerformance;
              messages.push("✓ Best-sellers recalculated using POS Data (net revenue after discounts)");
            }
          }
        } else if (posSheetName && isHistoricalYear) {
          messages.push(`ℹ️ ${year} is a past year, so the POS Data sheet was skipped — only the Monthly Sales summary was applied.`);
        }

        const customOrdersSheetName = findSheetName(wb.SheetNames, "custom orders");
        if (customOrdersSheetName) {
          const customOrders = parseCustomOrdersSheet(wb.Sheets[customOrdersSheetName]);
          if (customOrders) {
            fields.customOrders = customOrders;
            messages.push(`✓ Custom Orders sheet recognized (${customOrders.length} order line(s))`);
          }
        }
        const leadsSheetName = findSheetName(wb.SheetNames, "custom order leads") || findSheetName(wb.SheetNames, "leads");
        if (leadsSheetName) {
          const customLeads = parseCustomOrderLeadsSheet(wb.Sheets[leadsSheetName]);
          if (customLeads) {
            fields.customLeads = customLeads;
            messages.push(`✓ Custom Order Leads sheet recognized (${customLeads.length} lead(s))`);
          }
        }
        const marketEventsSheetName = findSheetName(wb.SheetNames, "market events");
        if (marketEventsSheetName) {
          const yearRecordsForEvents = (posResult && posResult.yearRecords) || [];
          const marketEvents = parseMarketEventDatesSheet(wb.Sheets[marketEventsSheetName], yearRecordsForEvents);
          if (marketEvents) {
            fields.marketEvents = marketEvents;
            messages.push(`✓ Market Events sheet recognized (${marketEvents.length} event day(s))`);
          }
        }
        // 별도 탭이 없으면, Monthly Sales 시트의 "Custom Order" 행에 달린 메모(구글 시트 "메모" 기능,
        // xlsx로 내보내면 쓰레드 코멘트로 저장됨)를 대신 읽어옴. "Customer / Category / Product / Qty
        // / Amount" 형식의 줄과 "LEAD: ..." 형식의 줄을 인식함 — 탭 방식과 메모 방식 둘 다 지원하며,
        // 탭이 있으면 탭을 우선함.
        if (!customOrdersSheetName || !leadsSheetName) {
          const memoResult = parseCustomOrderMemoSheet(wb.Sheets[salesSheetName]);
          if (memoResult && memoResult.debug && memoResult.debug.length > 0) {
            messages.push(`🔍 Custom Order memo 점검: ${memoResult.debug.join(" / ")}`);
          }
          if (memoResult) {
            if (!customOrdersSheetName && memoResult.orders.length > 0) {
              fields.customOrders = memoResult.orders;
              messages.push(`✓ Custom Order memo(s) recognized on the Monthly Sales sheet (${memoResult.orders.length} order line(s))`);
            }
            if (!leadsSheetName && memoResult.leads != null) {
              fields.customLeads = memoResult.leads;
              messages.push(memoResult.leads.length > 0
                ? `✓ Custom Order lead memo(s) recognized (${memoResult.leads.length} lead(s), from the most recent month's memo)`
                : "✓ Custom Order lead memo checked — no open leads in the most recent month's memo, so the lead list was cleared");
            }
            if (memoResult.unparsedCount > 0) {
              messages.push(`ℹ️ ${memoResult.unparsedCount} line(s) in the Custom Order memo didn't match the "Customer / Category / Product / Qty / Amount" format and were skipped (older free-text memos need to be rewritten in the new format to be picked up automatically).`);
            }
          }
        }

        bundles.push({ type: "sales", year: year || posYear, fields });
      } else if (dailySheetName) {
        const r = parseDailyReportSheet(wb.Sheets[dailySheetName]);
        const fields = {
          visits: r.visits, contacts: r.contacts, sold: r.sold, dowStats: r.dowStats,
          visitorOrigin: r.visitorOrigin, visitorType: r.visitorType, eventKeywords: r.eventKeywords,
        };
        if (r.operationStats) fields.operationStats = r.operationStats;
        const pivot = findVisitsPivotSheet(wb);
        if (pivot) {
          fields.visits = pivot.visits;
          fields.sold = pivot.sold;
          messages.push("✓ Pivot table summary found — used it for visits/units sold as the primary source (more reliable than day-by-day reclassification)");
        }
        messages.push(`✓ Daily report file recognized${r.year ? ` (${r.year}, ` : " ("}${r.monthsWithData} month(s) applied)`);
        messages.push(`🔍 Diagnostic: final August visits = ${(fields.visits[7] || 0).toLocaleString()} (all months: ${fields.visits.join(", ")})`);
        bundles.push({ type: "daily", year: r.year, fields });
      } else {
        messages.push(`✗ '${file.name}' — no recognizable sheet found. The sheet name must include 'Monthly Sales' or 'Daily Report'.`);
      }
    } catch (e) {
      messages.push(`✗ Error processing '${file.name}': ${e.message}`);
    }
  }
  return { bundles, messages };
}

/* ---------------------------------- component ---------------------------------- */
// The dashboard used to store these values in Korean. Any browser that already saved data
// (via localStorage) before the English update will keep replaying those old Korean strings
// on every load, since saved data normally takes priority over the (now-English) defaults.
// This table lets old saved records self-heal to English automatically, with no need to
// clear browser storage or re-upload anything.
const LEGACY_KO_LABELS = {
  category: {
    "관공서": "Government Office",
    "호텔·레스토랑": "Hotel/Restaurant",
    "호텔/레스토랑": "Hotel/Restaurant",
    "개인": "Individual",
    "기타": "Other",
  },
  method: {
    "현금 (Cash)": "Cash",
    "카드 (Card)": "Card",
    "MTN 모바일머니": "MTN Mobile Money",
    "Airtel 모바일머니": "Airtel Mobile Money",
    "은행 (Bank)": "Bank",
  },
};
function translateLegacyLabel(kind, value) {
  if (value == null) return value;
  const table = LEGACY_KO_LABELS[kind];
  return (table && table[value]) || value;
}
function normalizeCustomOrder(o) {
  return {
    id: o.id,
    groupId: o.groupId != null ? o.groupId : o.id,
    month: o.month != null ? o.month : 0,
    customer: o.customer != null ? o.customer : "",
    amount: o.amount != null ? o.amount : 0,
    note: o.note != null ? o.note : "",
    category: o.category !== undefined ? translateLegacyLabel("category", o.category) : null,
    product: o.product !== undefined ? o.product : null,
    qty: o.qty !== undefined ? o.qty : null,
  };
}

// 이벤트의 '월'은 항상 '날짜' 문자열에서 다시 계산해서 신뢰함 — 날짜만 입력하고 월을 안 맞춘 채
// 저장된 예전 데이터(예: 항상 1월로 기본값이 남아있는 경우)도 다음 로드 때 자동으로 바로잡힘
function normalizeMarketEvent(e) {
  const m = e.date ? String(e.date).match(/^\d{4}-(\d{1,2})-\d{1,2}/) : null;
  const derivedMonth = m ? Math.min(11, Math.max(0, parseInt(m[1], 10) - 1)) : e.month;
  return { ...e, month: derivedMonth != null ? derivedMonth : 0 };
}

// 새로 업로드한 파일이 일부 달만 담고 있어도(예: 최근 1개월치만) 이전에 이미 알고 있던 달의
// 데이터가 사라지지 않도록, 월별 배열은 "새 값이 0이 아니면 새 값, 아니면 기존 값 유지" 방식으로 병합
function mergeMonthlySeries(oldArr, newArr) {
  if (!Array.isArray(newArr)) return oldArr;
  if (!Array.isArray(oldArr)) return newArr;
  return MONTH_LABELS.map((_, i) => {
    const nv = Number(newArr[i]) || 0;
    const ov = Number(oldArr[i]) || 0;
    return nv > 0 ? nv : ov;
  });
}
function mergeMonthlyByKeyMap(oldObj, newObj) {
  const old = oldObj || {};
  const neu = newObj || {};
  const result = {};
  new Set([...Object.keys(old), ...Object.keys(neu)]).forEach((k) => {
    result[k] = mergeMonthlySeries(old[k], neu[k]);
  });
  return result;
}

// vendorMonthlyDetail은 {qty,gross,net,discount,count} 객체가 달마다 들어있는 구조라
// 위 mergeMonthlySeries로는 못 다뤄서, 달 단위로 "새 값이 있으면(qty나 net가 0보다 크면) 새 값,
// 아니면 기존 값 유지" 방식으로 따로 병합
function mergeVendorMonthlyDetail(oldDetail, newDetail) {
  const old = oldDetail || {};
  const neu = newDetail || {};
  const result = {};
  new Set([...Object.keys(old), ...Object.keys(neu)]).forEach((vendor) => {
    const oldArr = old[vendor];
    const newArr = neu[vendor];
    if (!Array.isArray(newArr)) { result[vendor] = oldArr; return; }
    if (!Array.isArray(oldArr)) { result[vendor] = newArr; return; }
    result[vendor] = MONTH_LABELS.map((_, i) => {
      const nv = newArr[i] || { qty: 0, gross: 0, net: 0, discount: 0, count: 0 };
      const ov = oldArr[i] || { qty: 0, gross: 0, net: 0, discount: 0, count: 0 };
      return (nv.qty > 0 || nv.net > 0) ? nv : ov;
    });
  });
  return result;
}

// customOrders/marketEvents는 구글 시트 쪽 소스(탭 또는 메모)가 "이번에 다시 확인한 달"만 담고
// 있을 수 있어(특히 메모 방식은 달마다 셀이 따로 있어, 아직 새 형식으로 안 옮긴 달은 이번 실행에
// 아예 안 잡힘), 그런 달까지 통째로 지워버리면 안 됨 -- 새로 읽어온 데이터에 포함된 "달"만 새
// 값으로 바꾸고, 새로 읽어온 데이터에 없는 달의 기존 주문/이벤트는 그대로 남겨둠.
function mergeRecordsByMonth(oldRecords, newRecords) {
  if (!Array.isArray(newRecords)) return oldRecords;
  if (!Array.isArray(oldRecords)) return newRecords;
  const newMonths = new Set(newRecords.map((r) => r.month));
  const preserved = oldRecords.filter((r) => !newMonths.has(r.month));
  return [...preserved, ...newRecords];
}

// These fields are entirely derived from processing an uploaded Excel file (or, now, the
// automated Google Sheet sync) -- nothing in the UI lets a person hand-edit them directly in
// a way that's meant to stick. So there's no reason a browser's old local save should ever win
// over a fresher automated upload for them; whichever side has the newer lastUploadedAt should
// simply provide these wholesale. customOrders, marketEvents, and customLeads moved into this
// list once the Custom Orders / Market Events / Custom Order Leads sheets were added to the
// automation -- the in-app "Edit" forms for these still work, but are now a local-only preview
// that the next automated sync will replace, same as the Data Management manual-edit table
// already behaves for visits/sold. (Fields with no sheet source at all -- productGroups,
// monthlyGoalUSD overrides, exchangeRate, consignment toggles, etc. -- are deliberately left out
// of this list so a person's own edits are never discarded just because the automation happened
// to run more recently.)
const AUTO_DERIVED_KEYS = [
  "dowStats", "visitorOrigin", "visitorType", "eventKeywords", "operationStats",
  "topProducts", "productPerformance", "allProducts", "paymentStats", "vendorStats",
  "customOrders", "marketEvents", "customLeads",
];
function mergeWithDefaults(defaults, loaded) {
  const merged = { ...defaults, ...loaded };
  // 얕은 병합만 하면 operationStats처럼 중첩된 객체에 나중에 추가된 하위 필드가
  // 예전에 저장된 데이터로 덮어써지면서 사라질 수 있어, 아래 필드는 한 겹 더 깊이 병합함
  const nestedKeys = ["operationStats", "priorYears", "itemVendorMap"];
  nestedKeys.forEach((k) => {
    if (defaults[k] && typeof defaults[k] === "object" && !Array.isArray(defaults[k])) {
      merged[k] = { ...defaults[k], ...(loaded && loaded[k] ? loaded[k] : {}) };
    }
  });
  // If the file's own baked-in data (from the latest automated run) is newer than whatever
  // this browser saved locally, the local save is stale for anything upload-derived -- use
  // the fresh defaults for those fields outright rather than the old cached snapshot.
  const defaultsAreNewer = !loaded || !loaded.lastUploadedAt || (defaults.lastUploadedAt && defaults.lastUploadedAt > loaded.lastUploadedAt);
  if (defaultsAreNewer) {
    AUTO_DERIVED_KEYS.forEach((k) => { merged[k] = defaults[k]; });
    merged.lastUploadedAt = defaults.lastUploadedAt;
  }
  // 연도가 바뀐 시점에 예전 저장값의 currentYear가 최신 자동화 값보다 앞서는 일은 없어야 함
  merged.currentYear = Math.max(defaults.currentYear || 0, (loaded && loaded.currentYear) || 0);
  // channels/vendorMonthly/posMonthlyQty/posInvoiceCounts는 월별 배열이라, 저장된 값 중 0인(비어있는) 달은
  // 기본 시드 데이터로 보완해서 예전 업로드가 일부 달만 담고 있었어도 이미 알던 달이 사라지지 않게 함
  if (loaded && loaded.channels) {
    const mergedChannels = {};
    const allKeys = new Set([...CHANNEL_KEYS, ...Object.keys(defaults.channels || {}), ...Object.keys(loaded.channels || {})]);
    allKeys.forEach((k) => { mergedChannels[k] = mergeMonthlySeries(defaults.channels[k], loaded.channels[k]); });
    merged.channels = mergedChannels;
  }
  if (defaults.extraChannels || (loaded && loaded.extraChannels)) {
    const byKey = {};
    (defaults.extraChannels || []).forEach((c) => { byKey[c.key] = c; });
    (loaded && loaded.extraChannels ? loaded.extraChannels : []).forEach((c) => { byKey[c.key] = c; });
    merged.extraChannels = Object.values(byKey);
  }
  merged.vendorMonthly = loaded && loaded.vendorMonthly
    ? mergeMonthlyByKeyMap(defaults.vendorMonthly, loaded.vendorMonthly)
    : defaults.vendorMonthly;
  merged.vendorMonthlyDetail = loaded && loaded.vendorMonthlyDetail
    ? mergeVendorMonthlyDetail(defaults.vendorMonthlyDetail, loaded.vendorMonthlyDetail)
    : defaults.vendorMonthlyDetail;
  if (loaded && loaded.posMonthlyQty) merged.posMonthlyQty = mergeMonthlySeries(defaults.posMonthlyQty, loaded.posMonthlyQty);
  if (loaded && loaded.posInvoiceCounts) merged.posInvoiceCounts = mergeMonthlySeries(defaults.posInvoiceCounts, loaded.posInvoiceCounts);
  // visits/contacts/sold are monthly arrays too (Daily Report), so they need the same
  // "prefer whichever side has a non-zero value" treatment -- otherwise a browser that has
  // ever saved data locally (via localStorage) freezes at whatever month it last saw, and a
  // newer month added by the daily automation run silently gets hidden behind the old save.
  if (loaded && loaded.visits) merged.visits = mergeMonthlySeries(defaults.visits, loaded.visits);
  if (loaded && loaded.contacts) merged.contacts = mergeMonthlySeries(defaults.contacts, loaded.contacts);
  if (loaded && loaded.sold) merged.sold = mergeMonthlySeries(defaults.sold, loaded.sold);
  // customOrders/marketEvents/customLeads are now sheet-derived (see AUTO_DERIVED_KEYS above),
  // so their freshness is already handled by the defaultsAreNewer block. No separate merge
  // needed here anymore -- the initial `{...defaults, ...loaded}` spread already provides the
  // right fallback on either side.
  // productGroups는, 저장된 목록에 없는 기본값 그룹만 추가해줌 (id 기준)
  if (loaded && Array.isArray(loaded.productGroups)) {
    const loadedGroupIds = new Set(loaded.productGroups.map((g) => g.id));
    const missingDefaultGroups = (defaults.productGroups || []).filter((g) => !loadedGroupIds.has(g.id));
    merged.productGroups = [...loaded.productGroups, ...missingDefaultGroups];
  } else {
    merged.productGroups = defaults.productGroups;
  }
  if (Array.isArray(merged.customOrders)) {
    merged.customOrders = merged.customOrders.map(normalizeCustomOrder);
  }
  if (Array.isArray(merged.marketEvents)) {
    merged.marketEvents = merged.marketEvents.map(normalizeMarketEvent);
  }
  // Self-heal old Korean payment-method labels the same way (see LEGACY_KO_LABELS above).
  if (Array.isArray(merged.paymentStats)) {
    merged.paymentStats = merged.paymentStats.map((p) => ({ ...p, method: translateLegacyLabel("method", p.method) }));
  }
  return merged;
}

export default function Dashboard() {
  const [data, setData] = useState(DEFAULT_DATA);
  const [ready, setReady] = useState(false);
  const [tab, setTab] = useState("overview");
  const [view, setView] = useState("monthly"); // monthly | quarterly | annual
  const [editMode, setEditMode] = useState(false);
  const [draft, setDraft] = useState(null);
  const [toast, setToast] = useState("");
  const [orderEditMode, setOrderEditMode] = useState(false);
  const [orderDraft, setOrderDraft] = useState(null);
  const [eventEditMode, setEventEditMode] = useState(false);
  const [eventDraft, setEventDraft] = useState(null);
  const [leadsEditMode, setLeadsEditMode] = useState(false);
  const [leadsDraft, setLeadsDraft] = useState(null);
  const [groupFormOpen, setGroupFormOpen] = useState(false);
  const [groupNameInput, setGroupNameInput] = useState("");
  const [groupMembersInput, setGroupMembersInput] = useState("");
  const [uploadBusy, setUploadBusy] = useState(false);
  const [uploadLog, setUploadLog] = useState([]);

  useEffect(() => {
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const loaded = JSON.parse(raw);
        setData(mergeWithDefaults(DEFAULT_DATA, loaded));
      }
    } catch (e) {
      /* no saved data yet, or storage unavailable — use the published sample data */
    }
    setReady(true);
  }, []);

  const showToast = (msg) => {
    setToast(msg);
    setTimeout(() => setToast(""), 2500);
  };

  async function persist(next) {
    setData(next);
    try {
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
      showToast("Saved to this browser (not reflected for other people)");
    } catch (e) {
      showToast("Save failed. Please try again.");
    }
  }

  async function toggleConsignmentVendor(vendorName) {
    const current = data.consignmentVendors || [];
    const next = current.includes(vendorName) ? current.filter((v) => v !== vendorName) : [...current, vendorName];
    await persist({ ...data, consignmentVendors: next });
  }

  async function handleFileUpload(fileList) {
    if (!fileList || fileList.length === 0) return;
    setUploadBusy(true);
    setUploadLog([]);
    try {
      const { bundles, messages } = await processExcelFiles(fileList, data.currentYear || 2026);
      if (bundles.length > 0) {
        const existingCurrentYear = data.currentYear || 2026;
        const detectedYears = bundles.map((b) => b.year).filter((y) => y != null);
        const newCurrentYear = detectedYears.length > 0 ? Math.max(existingCurrentYear, ...detectedYears) : existingCurrentYear;

        let nextTop = { ...data };
        let nextPriorYears = { ...(data.priorYears || {}) };
        let nextExtraChannels = [...(data.extraChannels || [])];

        // 더 최신 연도 파일이 들어와서 currentYear가 바뀌면, 기존 currentYear 데이터는 priorYears로 보존
        if (newCurrentYear !== existingCurrentYear) {
          nextPriorYears = { ...nextPriorYears, [existingCurrentYear]: extractYearFields(nextTop) };
        }

        bundles.forEach((b) => {
          const y = b.year || newCurrentYear;
          if (b.fields.newChannels) {
            b.fields.newChannels.forEach((nc) => {
              if (!nextExtraChannels.some((c) => c.key === nc.key)) {
                const color = EXTRA_CHANNEL_COLORS[nextExtraChannels.length % EXTRA_CHANNEL_COLORS.length];
                nextExtraChannels = [...nextExtraChannels, { key: nc.key, label: nc.label, color }];
              }
            });
          }
          if (y === newCurrentYear) {
            const mergedFields = { ...b.fields };
            delete mergedFields.newChannels;
            if (b.fields.channels) {
              const allKeys = new Set([...CHANNEL_KEYS, ...Object.keys(nextTop.channels || {}), ...Object.keys(b.fields.channels)]);
              const mergedChannels = {};
              allKeys.forEach((k) => {
                mergedChannels[k] = mergeMonthlySeries(nextTop.channels ? nextTop.channels[k] : null, b.fields.channels[k]);
              });
              mergedFields.channels = mergedChannels;
            }
            if (b.fields.vendorMonthly) mergedFields.vendorMonthly = mergeMonthlyByKeyMap(nextTop.vendorMonthly, b.fields.vendorMonthly);
            if (b.fields.vendorMonthlyDetail) mergedFields.vendorMonthlyDetail = mergeVendorMonthlyDetail(nextTop.vendorMonthlyDetail, b.fields.vendorMonthlyDetail);
            if (b.fields.itemVendorMap) mergedFields.itemVendorMap = { ...(nextTop.itemVendorMap || {}), ...b.fields.itemVendorMap };
            if (b.fields.posMonthlyQty) mergedFields.posMonthlyQty = mergeMonthlySeries(nextTop.posMonthlyQty, b.fields.posMonthlyQty);
            if (b.fields.posInvoiceCounts) mergedFields.posInvoiceCounts = mergeMonthlySeries(nextTop.posInvoiceCounts, b.fields.posInvoiceCounts);
            if (b.fields.visits) mergedFields.visits = mergeMonthlySeries(nextTop.visits, b.fields.visits);
            if (b.fields.contacts) mergedFields.contacts = mergeMonthlySeries(nextTop.contacts, b.fields.contacts);
            if (b.fields.sold) mergedFields.sold = mergeMonthlySeries(nextTop.sold, b.fields.sold);
            nextTop = { ...nextTop, ...mergedFields };
          } else {
            const prev = nextPriorYears[y] || {};
            const mergedFields = { ...b.fields };
            delete mergedFields.newChannels;
            if (b.fields.channels) {
              const allKeys = new Set([...CHANNEL_KEYS, ...Object.keys(prev.channels || {}), ...Object.keys(b.fields.channels)]);
              const mergedChannels = {};
              allKeys.forEach((k) => {
                mergedChannels[k] = mergeMonthlySeries(prev.channels ? prev.channels[k] : null, b.fields.channels[k]);
              });
              mergedFields.channels = mergedChannels;
            }
            nextPriorYears = { ...nextPriorYears, [y]: { ...prev, ...mergedFields } };
          }
        });

        const next = {
          ...nextTop, currentYear: newCurrentYear, priorYears: nextPriorYears,
          extraChannels: nextExtraChannels, lastUploadedAt: new Date().toISOString(),
        };
        await persist(next);
      }
      setUploadLog(messages);
    } catch (e) {
      setUploadLog([`✗ An error occurred while processing: ${e.message}`]);
    }
    setUploadBusy(false);
  }

  function startEdit() {
    setDraft(JSON.parse(JSON.stringify(data)));
    setEditMode(true);
  }
  function cancelEdit() {
    setDraft(null);
    setEditMode(false);
  }
  async function saveEdit() {
    await persist(draft);
    setEditMode(false);
    setDraft(null);
  }
  async function resetSample() {
    await persist(JSON.parse(JSON.stringify(DEFAULT_DATA)));
    showToast("Reset to sample data");
  }

  function updateDraftCell(kind, key, i, val) {
    const num = val === "" ? 0 : Number(val);
    setDraft((prev) => {
      const next = JSON.parse(JSON.stringify(prev));
      if (kind === "channel") next.channels[key][i] = isNaN(num) ? 0 : num;
      else next[kind][i] = isNaN(num) ? 0 : num;
      return next;
    });
  }

  function startEditOrders() {
    setOrderDraft(JSON.parse(JSON.stringify(data.customOrders)));
    setOrderEditMode(true);
  }
  function cancelEditOrders() {
    setOrderDraft(null);
    setOrderEditMode(false);
  }
  async function saveEditOrders() {
    await persist({ ...data, customOrders: orderDraft });
    setOrderEditMode(false);
    setOrderDraft(null);
  }
  async function reloadOrdersFromDefaults() {
    await persist({ ...data, customOrders: JSON.parse(JSON.stringify(DEFAULT_DATA.customOrders)) });
    setOrderEditMode(false);
    setOrderDraft(null);
  }
  function addOrderRow() {
    const newId = Date.now();
    setOrderDraft((prev) => [...prev, { id: newId, groupId: newId, month: 0, customer: "", amount: 0, note: "", category: null, product: "", qty: null }]);
  }
  function addProductLineToOrder(idx) {
    setOrderDraft((prev) => {
      const base = prev[idx];
      const newRow = {
        ...base, id: Date.now(), groupId: base.groupId ?? base.id,
        amount: 0, product: "", qty: null,
      };
      const next = [...prev];
      next.splice(idx + 1, 0, newRow);
      return next;
    });
  }
  function removeOrderRow(idx) {
    setOrderDraft((prev) => prev.filter((_, i) => i !== idx));
  }
  function updateOrderRow(idx, field, value) {
    setOrderDraft((prev) => {
      const next = [...prev];
      if (field === "amount" || field === "month") next[idx] = { ...next[idx], [field]: Number(value) || 0 };
      else if (field === "qty") next[idx] = { ...next[idx], qty: value === "" ? null : Number(value) };
      else next[idx] = { ...next[idx], [field]: value };
      return next;
    });
  }

  function startEditEvents() {
    setEventDraft(JSON.parse(JSON.stringify(data.marketEvents || [])));
    setEventEditMode(true);
  }
  function cancelEditEvents() {
    setEventDraft(null);
    setEventEditMode(false);
  }
  async function saveEditEvents() {
    await persist({ ...data, marketEvents: eventDraft.map(normalizeMarketEvent) });
    setEventEditMode(false);
    setEventDraft(null);
  }
  function addEventRow() {
    setEventDraft((prev) => [...prev, { id: Date.now(), date: "", month: 0, name: "", revenue: 0, qty: 0, invoices: 0 }]);
  }
  function removeEventRow(idx) {
    setEventDraft((prev) => prev.filter((_, i) => i !== idx));
  }
  function updateEventRow(idx, field, value) {
    setEventDraft((prev) => {
      const next = [...prev];
      if (field === "date") {
        // 날짜를 입력/수정하면 그래프 분류용 '월' 값도 자동으로 같이 계산해줌 (YYYY-MM-DD 형식 기준)
        const m = String(value).match(/^\d{4}-(\d{1,2})-\d{1,2}/);
        const derivedMonth = m ? Math.min(11, Math.max(0, parseInt(m[1], 10) - 1)) : next[idx].month;
        next[idx] = { ...next[idx], date: value, month: derivedMonth };
      } else if (["revenue", "qty", "invoices", "month"].includes(field)) {
        next[idx] = { ...next[idx], [field]: Number(value) || 0 };
      } else {
        next[idx] = { ...next[idx], [field]: value };
      }
      return next;
    });
  }

  function startEditLeads() {
    setLeadsDraft([...(data.customLeads || [])]);
    setLeadsEditMode(true);
  }
  function cancelEditLeads() {
    setLeadsDraft(null);
    setLeadsEditMode(false);
  }
  async function saveEditLeads() {
    await persist({ ...data, customLeads: leadsDraft.filter((l) => l.trim() !== "") });
    setLeadsEditMode(false);
    setLeadsDraft(null);
  }
  function addLead() {
    setLeadsDraft((prev) => [...prev, ""]);
  }
  function removeLead(idx) {
    setLeadsDraft((prev) => prev.filter((_, i) => i !== idx));
  }
  function updateLead(idx, value) {
    setLeadsDraft((prev) => {
      const next = [...prev];
      next[idx] = value;
      return next;
    });
  }

  async function addProductGroup() {
    const canonicalName = groupNameInput.trim();
    const members = groupMembersInput.split(",").map((s) => s.trim()).filter(Boolean);
    if (!canonicalName || members.length < 2) return;
    const nextGroups = [...(data.productGroups || []), { id: Date.now(), canonicalName, members }];
    await persist({ ...data, productGroups: nextGroups });
    setGroupNameInput("");
    setGroupMembersInput("");
    setGroupFormOpen(false);
  }
  async function removeProductGroup(id) {
    await persist({ ...data, productGroups: (data.productGroups || []).filter((g) => g.id !== id) });
  }

  const monthlyGoalUGX = data.monthlyGoalUSD * data.exchangeRate;

  /* ------------------------------ derived data ------------------------------ */
  const activeMonths = useMemo(() => activeMonthCount(data), [data]);
  const giftshopKeys = GIFTSHOP_KEYS;
  const channelMeta = CHANNEL_META;
  // 방문자/문의/판매수량은 일일 리포트 파일에서 오고, 매출과 별도로 업데이트될 수 있어
  // 매출 기준 activeMonths와 분리해서 자체적으로 '데이터가 있는 개월 수'를 계산
  const visitsActiveMonths = useMemo(
    () => Math.max(activeMonthsForArray(data.visits), activeMonthsForArray(data.contacts), activeMonthsForArray(data.sold)),
    [data]
  );

  // giftshop (all channels except custom/ODM)
  const giftshopMonthlyTotals = useMemo(
    () => MONTH_LABELS.map((_, i) => giftshopKeys.reduce((s, k) => s + (Number(data.channels[k][i]) || 0), 0)),
    [data]
  );
  const giftshopYTD = useMemo(
    () => giftshopMonthlyTotals.slice(0, activeMonths).reduce((a, b) => a + b, 0),
    [giftshopMonthlyTotals, activeMonths]
  );
  const giftshopYtdGoalUGX = monthlyGoalUGX * activeMonths;
  const giftshopAchievement = giftshopYtdGoalUGX > 0 ? giftshopYTD / giftshopYtdGoalUGX : 0;

  // Apoyo는 POS 매출 안에 이미 포함된 위탁판매분 — 실적 집계에서는 제외하고 참고용으로만 표시
  const consignmentVendors = data.consignmentVendors || [];
  const apoyoYTD = useMemo(
    () => data.channels.apoyo.slice(0, activeMonths).reduce((a, b) => a + Number(b || 0), 0),
    [data, activeMonths]
  );

  // 평균 거래단가 (POS 매출 ÷ 영수증 건수, giftshop 전체 - 위탁판매 포함 POS 거래 기준)
  const avgTransactionValue = useMemo(() => {
    const invoiceCounts = data.posInvoiceCounts || [];
    const posRevenue = data.channels.pos.slice(0, activeMonths).reduce((a, b) => a + Number(b || 0), 0);
    const invoiceTotal = invoiceCounts.slice(0, activeMonths).reduce((a, b) => a + Number(b || 0), 0);
    return invoiceTotal > 0 ? posRevenue / invoiceTotal : 0;
  }, [data, activeMonths]);

  const channelTotalsYTD = useMemo(() => {
    const obj = {};
    giftshopKeys.forEach((k) => {
      obj[k] = data.channels[k].slice(0, activeMonths).reduce((a, b) => a + Number(b || 0), 0);
    });
    return obj;
  }, [data, activeMonths]);

  const donutData = giftshopKeys.map((k) => ({
    name: channelMeta[k].label,
    value: channelTotalsYTD[k],
    color: channelMeta[k].color,
  })).filter((d) => d.value > 0);

  // trend chart data by view (giftshop channels only)
  // __labelTotal is an always-present marker field (1 whenever the month has any real
  // revenue, 0 otherwise) used purely to anchor the value-label placement below -- see the
  // comment on StackTotalLabel for why we can't rely on any one real channel for that.
  const trendData = useMemo(() => {
    if (view === "monthly") {
      return MONTH_LABELS.map((label, i) => {
        const row = { period: label };
        giftshopKeys.forEach((k) => (row[k] = data.channels[k][i] || 0));
        row.goal = monthlyGoalUGX;
        row.__labelTotal = giftshopKeys.reduce((s, k) => s + row[k], 0) > 0 ? 1 : 0;
        return row;
      });
    }
    if (view === "quarterly") {
      return QUARTER_LABELS.map((label, q) => {
        const row = { period: label };
        giftshopKeys.forEach((k) => {
          row[k] = [0, 1, 2].reduce((s, off) => s + (data.channels[k][q * 3 + off] || 0), 0);
        });
        row.goal = monthlyGoalUGX * 3;
        row.__labelTotal = giftshopKeys.reduce((s, k) => s + row[k], 0) > 0 ? 1 : 0;
        return row;
      });
    }
    const row = { period: `FY ${data.currentYear}` };
    giftshopKeys.forEach((k) => (row[k] = channelTotalsYTD[k]));
    row.goal = monthlyGoalUGX * 12;
    row.__labelTotal = giftshopKeys.reduce((s, k) => s + row[k], 0) > 0 ? 1 : 0;
    return [row];
  }, [view, data, monthlyGoalUGX, channelTotalsYTD, data.currentYear]);

  // ODM (custom order) — tracked separately with its own annual goal
  const odmMonthlyTotals = useMemo(() => MONTH_LABELS.map((_, i) => data.channels.custom[i] || 0), [data]);
  const odmYTD = useMemo(() => odmMonthlyTotals.slice(0, activeMonths).reduce((a, b) => a + b, 0), [odmMonthlyTotals, activeMonths]);
  const odmAnnualGoalUGX = data.odmAnnualGoalUGX || 0;
  const odmAchievement = odmAnnualGoalUGX > 0 ? odmYTD / odmAnnualGoalUGX : 0;
  const odmRemainingMonths = 12 - activeMonths;
  const odmRemainingGoal = Math.max(0, odmAnnualGoalUGX - odmYTD);
  const odmRequiredAvgRemaining = odmRemainingMonths > 0 ? odmRemainingGoal / odmRemainingMonths : 0;
  const odmTrendData = useMemo(
    () => MONTH_LABELS.map((label, i) => ({ period: label, amount: data.channels.custom[i] || 0, pace: odmAnnualGoalUGX / 12 })),
    [data, odmAnnualGoalUGX]
  );

  // 통합 매출 (B2C 기프트샵 + B2B 맞춤주문)
  const combinedYTD = giftshopYTD + odmYTD;
  const b2cShare = combinedYTD > 0 ? giftshopYTD / combinedYTD : 0;
  const b2bShare = combinedYTD > 0 ? odmYTD / combinedYTD : 0;
  const combinedTrendData = useMemo(() => {
    const withMarker = (row) => ({ ...row, __labelTotal: (row.b2c || 0) + (row.b2b || 0) > 0 ? 1 : 0 });
    if (view === "monthly") {
      return MONTH_LABELS.map((label, i) => withMarker({ period: label, b2c: giftshopMonthlyTotals[i] || 0, b2b: data.channels.custom[i] || 0 }));
    }
    if (view === "quarterly") {
      return QUARTER_LABELS.map((label, q) => withMarker({
        period: label,
        b2c: [0, 1, 2].reduce((s, off) => s + (giftshopMonthlyTotals[q * 3 + off] || 0), 0),
        b2b: [0, 1, 2].reduce((s, off) => s + (data.channels.custom[q * 3 + off] || 0), 0),
      }));
    }
    return [withMarker({ period: `FY ${data.currentYear}`, b2c: giftshopYTD, b2b: odmYTD })];
  }, [view, giftshopMonthlyTotals, data.channels.custom, giftshopYTD, odmYTD]);

  // 전년 대비 (YoY) — 같은 개월 수 기준으로 공정하게 비교
  const priorYearNum = data.currentYear - 1;
  const priorYearData = data.priorYears && data.priorYears[priorYearNum];
  const yoy = useMemo(() => {
    if (!priorYearData || !priorYearData.channels) return null;
    const priorGiftshopMonthly = MONTH_LABELS.map((_, i) => giftshopKeys.reduce((s, k) => s + (Number((priorYearData.channels[k] || [])[i]) || 0), 0));
    const priorOdmMonthly = MONTH_LABELS.map((_, i) => Number((priorYearData.channels.custom || [])[i]) || 0);
    const n = activeMonths;
    const priorGiftshopYTD = priorGiftshopMonthly.slice(0, n).reduce((a, b) => a + b, 0);
    const priorOdmYTD = priorOdmMonthly.slice(0, n).reduce((a, b) => a + b, 0);
    const priorCombinedYTD = priorGiftshopYTD + priorOdmYTD;
    const curCombinedYTD = giftshopYTD + odmYTD;
    const growth = (cur, prior) => (prior > 0 ? (cur - prior) / prior : null);
    return {
      year: priorYearNum,
      priorGiftshopYTD, priorOdmYTD, priorCombinedYTD,
      combinedGrowth: growth(curCombinedYTD, priorCombinedYTD),
      giftshopGrowth: growth(giftshopYTD, priorGiftshopYTD),
      odmGrowth: growth(odmYTD, priorOdmYTD),
    };
  }, [priorYearData, activeMonths, giftshopYTD, odmYTD, priorYearNum]);

  // 연도별 월별 추이 오버레이 (B2C/B2B 각각, 전년 vs 올해)
  const yearlyTrends = useMemo(() => {
    if (!priorYearData || !priorYearData.channels) return null;
    const priorGiftshopMonthly = MONTH_LABELS.map((_, i) => giftshopKeys.reduce((s, k) => s + (Number((priorYearData.channels[k] || [])[i]) || 0), 0));
    const priorOdmMonthly = MONTH_LABELS.map((_, i) => Number((priorYearData.channels.custom || [])[i]) || 0);
    const curKey = `y${data.currentYear}`;
    const priorKey = `y${priorYearNum}`;
    const b2c = MONTH_LABELS.map((label, i) => ({
      period: label, [priorKey]: priorGiftshopMonthly[i], [curKey]: i < activeMonths ? giftshopMonthlyTotals[i] : null,
    }));
    const b2b = MONTH_LABELS.map((label, i) => ({
      period: label, [priorKey]: priorOdmMonthly[i], [curKey]: i < activeMonths ? odmMonthlyTotals[i] : null,
    }));
    return { b2c, b2b, curKey, priorKey };
  }, [priorYearData, activeMonths, giftshopMonthlyTotals, odmMonthlyTotals, data.currentYear, priorYearNum]);

  // company-wide totals (used for visits/forecast tabs which are not split)
  const monthlyTotals = useMemo(() => MONTH_LABELS.map((_, i) => monthTotal(data, i)), [data]);
  const ytdRevenue = useMemo(
    () => monthlyTotals.slice(0, activeMonths).reduce((a, b) => a + b, 0),
    [monthlyTotals, activeMonths]
  );
  const ytdVisits = useMemo(() => data.visits.slice(0, visitsActiveMonths).reduce((a, b) => a + Number(b || 0), 0), [data, visitsActiveMonths]);
  const ytdContacts = useMemo(() => data.contacts.slice(0, visitsActiveMonths).reduce((a, b) => a + Number(b || 0), 0), [data, visitsActiveMonths]);
  // 판매 수량은 POS Data 시트가 가장 정확하다는 확인에 따라 이를 우선 사용, 없으면 Daily Report로 대체
  // 판매 수량은 베스트셀러 상품(Details Report)에서 이미 월별로 정리된 수량을 그대로 합산해서 사용함.
  // (POS Data 원본 거래를 다시 날짜별로 분류하는 방식은 일부 환경에서 불안정한 것이 확인되어 피하되,
  // 상품 목록과 같은 소스를 써서 베스트셀러/판매수량 숫자가 서로 어긋나지 않도록 함)
  const productBasedSoldQty = useMemo(() => {
    const arr = Array(12).fill(0);
    (data.allProducts || []).forEach((p) => { (p.qty || []).forEach((q, i) => { arr[i] += Number(q) || 0; }); });
    return arr;
  }, [data.allProducts]);
  const soldSeries = useMemo(() => (productBasedSoldQty.some((v) => v > 0) ? productBasedSoldQty : data.sold), [productBasedSoldQty, data.sold]);
  const soldIsFromPos = useMemo(() => productBasedSoldQty.some((v) => v > 0), [productBasedSoldQty]);
  const soldActiveMonths = useMemo(() => activeMonthsForArray(soldSeries), [soldSeries]);
  const ytdSold = useMemo(() => soldSeries.slice(0, soldActiveMonths).reduce((a, b) => a + Number(b || 0), 0), [soldSeries, soldActiveMonths]);
  const conversionRate = ytdVisits > 0 ? ytdContacts / ytdVisits : 0;

  // 향후 전망: B2C(기프트샵)와 B2B(맞춤주문)를 분리해서 각각 계산
  const b2cForecast = useMemo(
    () => computeForecast(giftshopMonthlyTotals, activeMonths, monthlyGoalUGX * 12),
    [giftshopMonthlyTotals, activeMonths, monthlyGoalUGX]
  );
  const b2bForecast = useMemo(
    () => computeForecast(odmMonthlyTotals, activeMonths, odmAnnualGoalUGX),
    [odmMonthlyTotals, activeMonths, odmAnnualGoalUGX]
  );

  // 연간 총액 비교: 전년(완료) 실적 vs 올해(실적 + 남은 기간 예상치)
  const annualCompare = useMemo(() => {
    if (!priorYearData || !priorYearData.channels) return null;
    const priorGiftshopFull = MONTH_LABELS.reduce((s, _, i) => s + giftshopKeys.reduce((s2, k) => s2 + (Number((priorYearData.channels[k] || [])[i]) || 0), 0), 0);
    const priorOdmFull = MONTH_LABELS.reduce((s, _, i) => s + (Number((priorYearData.channels.custom || [])[i]) || 0), 0);
    const b2cProjectedRemaining = b2cForecast ? Math.max(0, b2cForecast.projectedFlat - giftshopYTD) : 0;
    const b2bProjectedRemaining = b2bForecast ? Math.max(0, b2bForecast.projectedFlat - odmYTD) : 0;
    return {
      b2c: [
        { year: `${priorYearNum}`, actual: priorGiftshopFull, projected: 0 },
        { year: `${data.currentYear}`, actual: giftshopYTD, projected: b2cProjectedRemaining },
      ],
      b2b: [
        { year: `${priorYearNum}`, actual: priorOdmFull, projected: 0 },
        { year: `${data.currentYear}`, actual: odmYTD, projected: b2bProjectedRemaining },
      ],
    };
  }, [priorYearData, giftshopYTD, odmYTD, b2cForecast, b2bForecast, priorYearNum, data.currentYear]);

  if (!ready) {
    return (
      <div style={{ padding: 40, fontFamily: "Inter, sans-serif", color: COLORS.inkSoft }}>
        Loading data...
      </div>
    );
  }

  const tabs = [
    { id: "overview", label: "Overview", icon: TrendingUp },
    { id: "visits", label: "Visits & Conversion", icon: Users },
    { id: "custom", label: "Custom Orders", icon: Briefcase },
    { id: "vendor", label: "Vendor Analysis", icon: Factory },
    { id: "products", label: "Products & Payment", icon: ShoppingBag },
    { id: "forecast", label: "Forecast", icon: Sparkles },
    { id: "data", label: "Data Management", icon: Pencil },
  ];

  return (
    <div style={{ background: COLORS.bg, minHeight: "100%", fontFamily: "'Inter', sans-serif", color: COLORS.ink }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');
        * { box-sizing: border-box; }
        .sdc-scroll::-webkit-scrollbar { height: 8px; width: 8px; }
        .sdc-scroll::-webkit-scrollbar-thumb { background: ${COLORS.border}; border-radius: 4px; }
        input[type=number] { font-family: 'IBM Plex Mono', monospace; }
        input[type=number]::-webkit-inner-spin-button { opacity: 0.4; }
      `}</style>

      {/* header */}
      <div style={{ borderBottom: `1px solid ${COLORS.border}`, background: COLORS.surface, padding: "28px 32px 0" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", flexWrap: "wrap", gap: 16 }}>
          <div>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 6 }}>
              <div style={{ width: 10, height: 10, borderRadius: "50%", background: COLORS.ochre }} />
              <span style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, letterSpacing: 1.5, color: COLORS.inkFaint, textTransform: "uppercase" }}>
                SDC Gift Shop · Kampala
              </span>
            </div>
            <h1 style={{ fontFamily: "'Fraunces', serif", fontWeight: 700, fontSize: 34, margin: 0, color: COLORS.ink }}>
              {data.currentYear} Sales Dashboard
            </h1>
          </div>
          <div style={{ fontSize: 13, color: COLORS.inkSoft, textAlign: "right" }}>
            Data as of: {activeMonths > 0 ? `${data.currentYear} Jan - ${MONTH_LABELS[activeMonths - 1]} (${activeMonths} month${activeMonths > 1 ? "s" : ""})` : "No data entered"}
          </div>
        </div>
        <div style={{ display: "flex", gap: 4, marginTop: 22 }}>
          {tabs.map((t) => {
            const Icon = t.icon;
            const active = tab === t.id;
            return (
              <button
                key={t.id}
                onClick={() => {
                  setTab(t.id);
                  if (t.id !== "data") { setEditMode(false); setDraft(null); }
                  if (t.id !== "custom") { setOrderEditMode(false); setOrderDraft(null); }
                }}
                style={{
                  display: "flex", alignItems: "center", gap: 7, padding: "10px 16px",
                  background: "transparent", border: "none", cursor: "pointer",
                  fontFamily: "'Inter', sans-serif", fontSize: 14, fontWeight: 600,
                  color: active ? COLORS.teal : COLORS.inkFaint,
                  borderBottom: active ? `2.5px solid ${COLORS.teal}` : "2.5px solid transparent",
                  marginBottom: -1,
                }}
              >
                <Icon size={16} />
                <span>{t.label}</span>
              </button>
            );
          })}
        </div>
      </div>

      <div style={{ padding: 32, maxWidth: 1240, margin: "0 auto" }}>
        {tab === "overview" && (
          <OverviewTab
            giftshopYTD={giftshopYTD} giftshopAchievement={giftshopAchievement} monthlyGoalUGX={monthlyGoalUGX}
            ytdVisits={ytdVisits} conversionRate={conversionRate} view={view} setView={setView} trendData={trendData}
            donutData={donutData} activeMonths={activeMonths} visitsActiveMonths={visitsActiveMonths} apoyoYTD={apoyoYTD} avgTransactionValue={avgTransactionValue}
            consignmentVendors={consignmentVendors}
            odmYTD={odmYTD} odmAchievement={odmAchievement} odmAnnualGoalUGX={odmAnnualGoalUGX}
            odmRemainingGoal={odmRemainingGoal} odmRequiredAvgRemaining={odmRequiredAvgRemaining}
            odmRemainingMonths={odmRemainingMonths} odmTrendData={odmTrendData} marketEvents={data.marketEvents}
            combinedYTD={combinedYTD} b2cShare={b2cShare} b2bShare={b2bShare} combinedTrendData={combinedTrendData} yoy={yoy}
            giftshopKeys={giftshopKeys} channelMeta={channelMeta}
            yearlyTrends={yearlyTrends} priorYearNum={priorYearNum} currentYear={data.currentYear} annualCompare={annualCompare}
          />
        )}
        {tab === "visits" && (
          <VisitsTab
            data={data} activeMonths={visitsActiveMonths} ytdSold={ytdSold} soldSeries={soldSeries} soldIsFromPos={soldIsFromPos}
            eventEditMode={eventEditMode} eventDraft={eventDraft} startEditEvents={startEditEvents} cancelEditEvents={cancelEditEvents}
            saveEditEvents={saveEditEvents} addEventRow={addEventRow} removeEventRow={removeEventRow} updateEventRow={updateEventRow}
          />
        )}
        {tab === "custom" && (
          <CustomOrderTab
            data={data} editOrders={orderEditMode} orderDraft={orderDraft}
            startEditOrders={startEditOrders} cancelEditOrders={cancelEditOrders} saveEditOrders={saveEditOrders}
            addOrderRow={addOrderRow} removeOrderRow={removeOrderRow} updateOrderRow={updateOrderRow}
            addProductLineToOrder={addProductLineToOrder} reloadOrdersFromDefaults={reloadOrdersFromDefaults}
            editLeads={leadsEditMode} leadsDraft={leadsDraft} startEditLeads={startEditLeads}
            cancelEditLeads={cancelEditLeads} saveEditLeads={saveEditLeads} addLead={addLead}
            removeLead={removeLead} updateLead={updateLead}
          />
        )}
        {tab === "vendor" && <VendorTab data={data} toggleConsignmentVendor={toggleConsignmentVendor} />}
        {tab === "products" && (
          <ProductsTab
            data={data} groupFormOpen={groupFormOpen} setGroupFormOpen={setGroupFormOpen}
            groupNameInput={groupNameInput} setGroupNameInput={setGroupNameInput}
            groupMembersInput={groupMembersInput} setGroupMembersInput={setGroupMembersInput}
            addProductGroup={addProductGroup} removeProductGroup={removeProductGroup}
          />
        )}
        {tab === "forecast" && <ForecastTab b2cForecast={b2cForecast} b2bForecast={b2bForecast} data={data} activeMonths={activeMonths} />}
        {tab === "data" && (
          <DataTab
            data={data} editMode={editMode} draft={draft} startEdit={startEdit} cancelEdit={cancelEdit}
            saveEdit={saveEdit} resetSample={resetSample} updateDraftCell={updateDraftCell} setDraft={setDraft}
            onUploadFiles={handleFileUpload} uploadBusy={uploadBusy} uploadLog={uploadLog}
          />
        )}
      </div>

      {toast && (
        <div style={{
          position: "fixed", bottom: 24, left: "50%", transform: "translateX(-50%)",
          background: COLORS.ink, color: "#fff", padding: "10px 20px", borderRadius: 8,
          fontSize: 13, fontWeight: 500, boxShadow: "0 6px 20px rgba(0,0,0,0.2)", zIndex: 50,
        }}>
          {toast}
        </div>
      )}
    </div>
  );
}

/* ---------------------------------- shared bits ---------------------------------- */
function Card({ children, style }) {
  return (
    <div style={{
      background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 14,
      padding: 22, ...style,
    }}>
      {children}
    </div>
  );
}
function SectionTitle({ children, sub }) {
  return (
    <div style={{ marginBottom: 16 }}>
      <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 19, fontWeight: 600, margin: 0, color: COLORS.ink }}>{children}</h2>
      {sub && <p style={{ fontSize: 13, color: COLORS.inkFaint, margin: "4px 0 0" }}>{sub}</p>}
    </div>
  );
}
function KpiCard({ icon: Icon, label, value, accent, note }) {
  return (
    <Card style={{ flex: 1, minWidth: 200 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
        <div style={{ width: 30, height: 30, borderRadius: 8, background: accent + "22", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <Icon size={15} color={accent} />
        </div>
        <span style={{ fontSize: 12.5, color: COLORS.inkSoft, fontWeight: 600 }}>{label}</span>
      </div>
      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 24, fontWeight: 600, color: COLORS.ink }}>{value}</div>
      {note && <div style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 4 }}>{note}</div>}
    </Card>
  );
}
function ViewToggle({ view, setView }) {
  const opts = [["monthly", "Monthly"], ["quarterly", "Quarterly"], ["annual", "Annual"]];
  return (
    <div style={{ display: "flex", background: COLORS.bg, border: `1px solid ${COLORS.border}`, borderRadius: 9, padding: 3 }}>
      {opts.map(([id, label]) => (
        <button key={id} onClick={() => setView(id)} style={{
          padding: "6px 14px", borderRadius: 7, border: "none", cursor: "pointer",
          fontSize: 12.5, fontWeight: 600, fontFamily: "'Inter', sans-serif",
          background: view === id ? COLORS.teal : "transparent",
          color: view === id ? "#fff" : COLORS.inkSoft,
        }}>{label}</button>
      ))}
    </div>
  );
}
function MonthPicker({ month, setMonth }) {
  return (
    <div className="sdc-scroll" style={{ display: "flex", gap: 4, overflowX: "auto", padding: 3, background: COLORS.bg, borderRadius: 9, border: `1px solid ${COLORS.border}` }}>
      <button onClick={() => setMonth(null)} style={{
        padding: "6px 12px", borderRadius: 7, border: "none", cursor: "pointer", whiteSpace: "nowrap",
        fontSize: 12.5, fontWeight: 600, fontFamily: "'Inter', sans-serif",
        background: month === null ? COLORS.teal : "transparent", color: month === null ? "#fff" : COLORS.inkSoft,
      }}>All (Cumulative)</button>
      {MONTH_LABELS.map((label, i) => (
        <button key={label} onClick={() => setMonth(i)} style={{
          padding: "6px 12px", borderRadius: 7, border: "none", cursor: "pointer", whiteSpace: "nowrap",
          fontSize: 12.5, fontWeight: 600, fontFamily: "'Inter', sans-serif",
          background: month === i ? COLORS.teal : "transparent", color: month === i ? "#fff" : COLORS.inkSoft,
        }}>{label}</button>
      ))}
    </div>
  );
}
function tooltipStyle() {
  return { background: "#fff", border: `1px solid ${COLORS.border}`, borderRadius: 8, fontSize: 12.5, fontFamily: "'Inter', sans-serif" };
}

/* ---------------------------------- overview tab ---------------------------------- */
function TrendTick({ x, y, payload, eventMonthLabels }) {
  const has = eventMonthLabels && eventMonthLabels.has(payload.value);
  return (
    <text x={x} y={y + 16} textAnchor="middle" fontSize={12} fill={has ? COLORS.clay : COLORS.inkSoft} fontWeight={has ? 700 : 400}>
      {payload.value}{has ? " ★" : ""}
    </text>
  );
}

// Recharts' LabelList does not create an entry for a bar segment whose own value is 0
// (no rectangle is rendered for a zero-height segment), so the "index" it hands to a custom
// content renderer counts only the NON-ZERO segments of that particular series -- it silently
// skips zero months instead of leaving a gap. If we index straight into the full, unfiltered
// month array with that count, every month after the first zero ends up reading the wrong
// month's data (everything shifts left by one for each zero encountered so far). Mirroring
// that same skip here (filtering rows by the anchor series before indexing) keeps our lookup
// in sync with whatever position Recharts is actually rendering.
function StackTotalLabel({ x, y, width, index, rows, keys, anchorKey }) {
  const anchor = anchorKey || keys[keys.length - 1];
  const filteredRows = (rows || []).filter((r) => Number(r[anchor]) > 0);
  const row = filteredRows[index];
  if (!row) return null;
  const total = keys.reduce((s, k) => s + (Number(row[k]) || 0), 0);
  if (total <= 0) return null;
  return (
    <text x={x + width / 2} y={y - 8} textAnchor="middle" fontSize={11.5} fontWeight={700} fill={COLORS.ink}>
      {fmtCompact(total)}
    </text>
  );
}

// 단일 막대(스택 아닌) 차트용 값 라벨 — 0은 표시하지 않음, 숫자 포맷터(fmtCompact 등)를 선택 가능
function BarValueLabel({ x, y, width, value, formatter }) {
  const v = Number(value) || 0;
  if (v <= 0) return null;
  return (
    <text x={x + width / 2} y={y - 8} textAnchor="middle" fontSize={11.5} fontWeight={700} fill={COLORS.ink}>
      {(formatter || fmtCompact)(v)}
    </text>
  );
}

// 선그래프용 값 라벨 — 선/다른 라벨과 겹쳐도 읽히도록 흰 배경 칩을 깔고, dy로 위/아래 방향을 벌려줌
function LineValueLabel({ x, y, value, formatter, color, dy }) {
  const v = Number(value);
  if (!v || v <= 0) return null;
  const text = (formatter || fmtCompact)(v);
  const w = text.length * 6.4 + 10;
  const cy = y + (dy || 0);
  return (
    <g>
      <rect x={x - w / 2} y={cy - 11} width={w} height={16} rx={4} fill="#fff" fillOpacity={0.92} stroke={COLORS.border} strokeWidth={0.5} />
      <text x={x} y={cy + 1} textAnchor="middle" fontSize={10.5} fontWeight={700} fill={color}>{text}</text>
    </g>
  );
}


function OverviewTab({
  giftshopYTD, giftshopAchievement, monthlyGoalUGX, ytdVisits, conversionRate, view, setView, trendData, donutData, activeMonths, visitsActiveMonths, apoyoYTD, avgTransactionValue, consignmentVendors,
  odmYTD, odmAchievement, odmAnnualGoalUGX, odmRemainingGoal, odmRequiredAvgRemaining, odmRemainingMonths, odmTrendData, marketEvents,
  combinedYTD, b2cShare, b2bShare, combinedTrendData, yoy, yearlyTrends, priorYearNum, currentYear, annualCompare, giftshopKeys, channelMeta,
}) {
  const eventMonthLabels = new Set((marketEvents || []).map((e) => MONTH_LABELS[e.month]));
  const eventsByMonthLabel = {};
  (marketEvents || []).forEach((e) => {
    const label = MONTH_LABELS[e.month];
    if (!eventsByMonthLabel[label]) eventsByMonthLabel[label] = [];
    eventsByMonthLabel[label].push(e.name);
  });
  const combinedDonutData = [
    { name: "B2C · Gift Shop", value: giftshopYTD, color: COLORS.teal },
    { name: "B2B · Custom Orders", value: odmYTD, color: COLORS.ochre },
  ].filter((d) => d.value > 0);
  return (
    <div>
      {/* ---------------- combined B2C+B2B section ---------------- */}
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
        <div style={{ width: 10, height: 10, borderRadius: 3, background: COLORS.plum }} />
        <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 22, fontWeight: 700, margin: 0, color: COLORS.ink }}>Combined Revenue (B2C + B2B)</h2>
        <span style={{ fontSize: 11, color: COLORS.inkFaint, marginLeft: 2 }}>Combined Revenue</span>
        <span style={{ fontSize: 12.5, color: COLORS.inkFaint }}>(Gift Shop + Custom Orders combined)</span>
      </div>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 24 }}>
        <KpiCard icon={TrendingUp} label="Combined YTD Revenue" value={fmtCompact(combinedYTD) + " UGX"} accent={COLORS.plum} note={fmtUGX(combinedYTD)} />
        <KpiCard icon={ShoppingBag} label="B2C Share (Gift Shop)" value={fmtPct(b2cShare)} accent={COLORS.teal} note={fmtCompact(giftshopYTD) + " UGX"} />
        <KpiCard icon={Briefcase} label="B2B Share (Custom Orders)" value={fmtPct(b2bShare)} accent={COLORS.ochre} note={fmtCompact(odmYTD) + " UGX"} />
        {yoy && (
          <KpiCard
            icon={TrendingUp} label={`vs Same Period Last Year (${yoy.year})`}
            value={yoy.combinedGrowth === null ? "—" : `${yoy.combinedGrowth >= 0 ? "+" : ""}${(yoy.combinedGrowth * 100).toFixed(1)}%`}
            accent={yoy.combinedGrowth >= 0 ? COLORS.teal : COLORS.clay}
            note={`Compared over the same ${activeMonths} month(s)`}
          />
        )}
      </div>
      {yoy && (
        <p style={{ fontSize: 12.5, color: COLORS.inkFaint, marginTop: -10, marginBottom: 24 }}>
          <Info size={12} style={{ verticalAlign: -1, marginRight: 4 }} />
          vs same period {yoy.year} (same {activeMonths} months) — B2C {yoy.giftshopGrowth === null ? "—" : `${yoy.giftshopGrowth >= 0 ? "+" : ""}${(yoy.giftshopGrowth * 100).toFixed(1)}%`},
          B2B {yoy.odmGrowth === null ? "—" : `${yoy.odmGrowth >= 0 ? "+" : ""}${(yoy.odmGrowth * 100).toFixed(1)}%`}
        </p>
      )}

      <div style={{ display: "flex", gap: 20, flexWrap: "wrap", marginBottom: 40 }}>
        <Card style={{ flex: "2 1 480px" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, flexWrap: "wrap", gap: 12 }}>
            <SectionTitle sub="Total revenue combining B2C (Gift Shop) and B2B (Custom Orders), with monthly split">Combined Revenue Trend</SectionTitle>
            <ViewToggle view={view} setView={setView} />
          </div>
          <ResponsiveContainer width="100%" height={320}>
            <ComposedChart data={combinedTrendData} margin={{ top: 24, right: 10, left: 0, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} vertical={false} />
              <XAxis dataKey="period" tick={{ fontSize: 12, fill: COLORS.inkSoft }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
              <YAxis tickFormatter={fmtCompact} tick={{ fontSize: 11, fill: COLORS.inkFaint }} axisLine={false} tickLine={false} width={50} />
              <Tooltip contentStyle={tooltipStyle()} formatter={(v, name) => [fmtUGX(v), name === "b2c" ? "B2C · Gift Shop" : "B2B · Custom Orders"]} />
              <Legend formatter={(v) => (v === "b2c" ? "B2C · Gift Shop" : "B2B · Custom Orders")} wrapperStyle={{ fontSize: 12 }} />
              <Bar dataKey="b2c" stackId="combined" fill={COLORS.teal} radius={0} />
              <Bar dataKey="b2b" stackId="combined" fill={COLORS.ochre} radius={0} />
              <Bar dataKey="__labelTotal" stackId="combined" fill="transparent" radius={0} isAnimationActive={false}>
                <LabelList dataKey="__labelTotal" content={(props) => <StackTotalLabel {...props} rows={combinedTrendData} keys={["b2c", "b2b"]} anchorKey="__labelTotal" />} />
              </Bar>
            </ComposedChart>
          </ResponsiveContainer>
        </Card>
        <Card style={{ flex: "1 1 260px" }}>
          <SectionTitle sub="B2C vs B2B revenue share, cumulative">B2C / B2B Share</SectionTitle>
          <ResponsiveContainer width="100%" height={240}>
            <PieChart>
              <Pie data={combinedDonutData} dataKey="value" nameKey="name" innerRadius={54} outerRadius={88} paddingAngle={3}>
                {combinedDonutData.map((d, i) => <Cell key={i} fill={d.color} />)}
              </Pie>
              <Tooltip contentStyle={tooltipStyle()} formatter={(v) => fmtUGX(v)} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
            </PieChart>
          </ResponsiveContainer>
        </Card>
      </div>

      {yearlyTrends && (
        <>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
            <div style={{ width: 10, height: 10, borderRadius: 3, background: COLORS.slate }} />
            <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 22, fontWeight: 700, margin: 0, color: COLORS.ink }}>Year-over-Year Trend</h2>
            <span style={{ fontSize: 11, color: COLORS.inkFaint, marginLeft: 2 }}>Year-over-Year Trend</span>
            <span style={{ fontSize: 12.5, color: COLORS.inkFaint }}>({priorYearNum} vs {currentYear} · same month comparison)</span>
          </div>
          <div style={{ display: "flex", gap: 20, flexWrap: "wrap", marginBottom: 40 }}>
            <Card style={{ flex: "1 1 420px" }}>
              <SectionTitle sub="Monthly B2C (Gift Shop) revenue comparison">B2C Year-over-Year</SectionTitle>
              <ResponsiveContainer width="100%" height={340}>
                <ComposedChart data={yearlyTrends.b2c} margin={{ top: 30, right: 10, left: 0, bottom: 24 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} vertical={false} />
                  <XAxis dataKey="period" tick={{ fontSize: 12, fill: COLORS.inkSoft }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
                  <YAxis tickFormatter={fmtCompact} tick={{ fontSize: 11, fill: COLORS.inkFaint }} axisLine={false} tickLine={false} width={50} />
                  <Tooltip contentStyle={tooltipStyle()} formatter={(v) => fmtUGX(v)} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Line dataKey={yearlyTrends.priorKey} name={`${priorYearNum}`} stroke={COLORS.inkFaint} strokeWidth={2} strokeDasharray="4 3" dot={{ r: 3 }}>
                    <LabelList dataKey={yearlyTrends.priorKey} content={(p) => <LineValueLabel {...p} color={COLORS.inkFaint} dy={22} />} />
                  </Line>
                  <Line dataKey={yearlyTrends.curKey} name={`${currentYear}`} stroke={COLORS.teal} strokeWidth={2.5} dot={{ r: 3 }} connectNulls={false}>
                    <LabelList dataKey={yearlyTrends.curKey} content={(p) => <LineValueLabel {...p} color={COLORS.teal} dy={-20} />} />
                  </Line>
                </ComposedChart>
              </ResponsiveContainer>
            </Card>
            <Card style={{ flex: "1 1 420px" }}>
              <SectionTitle sub="Monthly B2B (Custom Orders) revenue comparison">B2B Year-over-Year</SectionTitle>
              <ResponsiveContainer width="100%" height={340}>
                <ComposedChart data={yearlyTrends.b2b} margin={{ top: 30, right: 10, left: 0, bottom: 24 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} vertical={false} />
                  <XAxis dataKey="period" tick={{ fontSize: 12, fill: COLORS.inkSoft }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
                  <YAxis tickFormatter={fmtCompact} tick={{ fontSize: 11, fill: COLORS.inkFaint }} axisLine={false} tickLine={false} width={50} />
                  <Tooltip contentStyle={tooltipStyle()} formatter={(v) => fmtUGX(v)} />
                  <Legend wrapperStyle={{ fontSize: 12 }} />
                  <Line dataKey={yearlyTrends.priorKey} name={`${priorYearNum}`} stroke={COLORS.inkFaint} strokeWidth={2} strokeDasharray="4 3" dot={{ r: 3 }}>
                    <LabelList dataKey={yearlyTrends.priorKey} content={(p) => <LineValueLabel {...p} color={COLORS.inkFaint} dy={22} />} />
                  </Line>
                  <Line dataKey={yearlyTrends.curKey} name={`${currentYear}`} stroke={COLORS.ochre} strokeWidth={2.5} dot={{ r: 3 }} connectNulls={false}>
                    <LabelList dataKey={yearlyTrends.curKey} content={(p) => <LineValueLabel {...p} color={COLORS.ochre} dy={-20} />} />
                  </Line>
                </ComposedChart>
              </ResponsiveContainer>
            </Card>
          </div>

          {annualCompare && (
            <div style={{ display: "flex", gap: 20, flexWrap: "wrap", marginBottom: 40 }}>
              <Card style={{ flex: "1 1 420px" }}>
                <SectionTitle sub={`${priorYearNum} actual vs ${currentYear} actual (solid) + projected remainder (light)`}>B2C Annual Total Comparison</SectionTitle>
                <ResponsiveContainer width="100%" height={240}>
                  <ComposedChart data={annualCompare.b2c} margin={{ top: 20, right: 10, left: 0, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} vertical={false} />
                    <XAxis dataKey="year" tick={{ fontSize: 12, fill: COLORS.inkSoft }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
                    <YAxis tickFormatter={fmtCompact} tick={{ fontSize: 11, fill: COLORS.inkFaint }} axisLine={false} tickLine={false} width={50} />
                    <Tooltip contentStyle={tooltipStyle()} formatter={(v, name) => [fmtUGX(v), name === "actual" ? "Actual" : "Projected (remaining)"]} />
                    <Legend formatter={(v) => (v === "actual" ? "Actual" : "Projected (remaining)")} wrapperStyle={{ fontSize: 12 }} />
                    <Bar dataKey="actual" stackId="a" fill={COLORS.teal} radius={0}>
                      <LabelList dataKey="actual" content={(props) => <StackTotalLabel {...props} rows={annualCompare.b2c} keys={["actual", "projected"]} anchorKey="actual" />} />
                    </Bar>
                    <Bar dataKey="projected" stackId="a" fill={COLORS.tealSoft} radius={[4, 4, 0, 0]} />
                  </ComposedChart>
                </ResponsiveContainer>
              </Card>
              <Card style={{ flex: "1 1 420px" }}>
                <SectionTitle sub={`${priorYearNum} actual vs ${currentYear} actual (solid) + projected remainder (light)`}>B2B Annual Total Comparison</SectionTitle>
                <ResponsiveContainer width="100%" height={240}>
                  <ComposedChart data={annualCompare.b2b} margin={{ top: 20, right: 10, left: 0, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} vertical={false} />
                    <XAxis dataKey="year" tick={{ fontSize: 12, fill: COLORS.inkSoft }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
                    <YAxis tickFormatter={fmtCompact} tick={{ fontSize: 11, fill: COLORS.inkFaint }} axisLine={false} tickLine={false} width={50} />
                    <Tooltip contentStyle={tooltipStyle()} formatter={(v, name) => [fmtUGX(v), name === "actual" ? "Actual" : "Projected (remaining)"]} />
                    <Legend formatter={(v) => (v === "actual" ? "Actual" : "Projected (remaining)")} wrapperStyle={{ fontSize: 12 }} />
                    <Bar dataKey="actual" stackId="a" fill={COLORS.ochre} radius={0}>
                      <LabelList dataKey="actual" content={(props) => <StackTotalLabel {...props} rows={annualCompare.b2b} keys={["actual", "projected"]} anchorKey="actual" />} />
                    </Bar>
                    <Bar dataKey="projected" stackId="a" fill={COLORS.ochreSoft} radius={[4, 4, 0, 0]} />
                  </ComposedChart>
                </ResponsiveContainer>
              </Card>
            </div>
          )}
        </>
      )}

      {/* ---------------- giftshop section ---------------- */}
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
        <div style={{ width: 10, height: 10, borderRadius: 3, background: COLORS.teal }} />
        <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 22, fontWeight: 700, margin: 0, color: COLORS.ink }}>Gift Shop Revenue</h2>
        <span style={{ fontSize: 11, color: COLORS.inkFaint, marginLeft: 2 }}>Gift Shop Revenue</span>
        <span style={{ fontSize: 12.5, color: COLORS.inkFaint }}>(excludes Custom Orders · POS/Apoyo/BOLD/Artisan/Yujo)</span>
      </div>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 24 }}>
        <KpiCard icon={TrendingUp} label="YTD Revenue" value={fmtCompact(giftshopYTD) + " UGX"} accent={COLORS.teal} note={fmtUGX(giftshopYTD)} />
        <KpiCard icon={Target} label="Goal Achievement" value={fmtPct(giftshopAchievement)} accent={giftshopAchievement >= 1 ? COLORS.teal : COLORS.clay} note={`Based on monthly goal ${fmtUGX(monthlyGoalUGX)}`} />
        <KpiCard icon={Users} label="Total Visitors" value={fmtNum(ytdVisits) + ""} accent={COLORS.ochre} note={`${visitsActiveMonths}-month total`} />
        <KpiCard icon={CreditCard} label="Avg. Transaction Value" value={fmtCompact(avgTransactionValue) + " UGX"} accent={COLORS.plum} note="POS revenue ÷ receipt count" />
      </div>

      <Card style={{ marginBottom: 24 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4, flexWrap: "wrap", gap: 12 }}>
          <SectionTitle sub="Revenue by store channel excluding custom orders, vs. monthly goal (★ = a month with a market/community event)">Gift Shop Revenue Trend</SectionTitle>
          <ViewToggle view={view} setView={setView} />
        </div>
        <ResponsiveContainer width="100%" height={360}>
          <ComposedChart data={trendData} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} vertical={false} />
            <XAxis dataKey="period" tick={(props) => <TrendTick {...props} eventMonthLabels={eventMonthLabels} />} axisLine={{ stroke: COLORS.border }} tickLine={false} />
            <YAxis tickFormatter={fmtCompact} tick={{ fontSize: 11, fill: COLORS.inkFaint }} axisLine={false} tickLine={false} width={50} />
            <Tooltip contentStyle={tooltipStyle()} formatter={(v, name) => [fmtUGX(v), channelMeta[name]?.label || name]} />
            <Legend formatter={(v) => channelMeta[v]?.label || v} wrapperStyle={{ fontSize: 12 }} />
            {giftshopKeys.map((k) => (
              <Bar key={k} dataKey={k} stackId="rev" fill={channelMeta[k].color} radius={0} />
            ))}
            <Bar dataKey="__labelTotal" stackId="rev" fill="transparent" radius={0} isAnimationActive={false}>
              <LabelList dataKey="__labelTotal" content={(props) => <StackTotalLabel {...props} rows={trendData} keys={giftshopKeys} anchorKey="__labelTotal" />} />
            </Bar>
            <Line dataKey="goal" stroke={COLORS.clay} strokeWidth={2} strokeDasharray="5 4" dot={false} name="Monthly Goal" />
          </ComposedChart>
        </ResponsiveContainer>
        {view === "monthly" && Object.keys(eventsByMonthLabel).length > 0 && (
          <p style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 10 }}>
            ★ {Object.entries(eventsByMonthLabel).map(([m, names]) => `${m}: ${names.join(", ")}`).join(" · ")}
          </p>
        )}
        {apoyoYTD > 0 && (
          <p style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 10 }}>
            <Info size={12} style={{ verticalAlign: -1, marginRight: 4 }} />
            Gift shop revenue (POS) includes {fmtUGX(apoyoYTD)} in Apoyo consignment sales. It's included in revenue as-is, but part of it is later settled to Apoyo, so SDC's actual take may differ.
          </p>
        )}
      </Card>

      <div style={{ display: "flex", gap: 20, flexWrap: "wrap", marginBottom: 40 }}>
        <Card style={{ flex: "1 1 340px" }}>
          <SectionTitle sub="Revenue share by channel, cumulative">Channel Revenue Mix</SectionTitle>
          <ResponsiveContainer width="100%" height={280}>
            <PieChart>
              <Pie data={donutData} dataKey="value" nameKey="name" innerRadius={62} outerRadius={98} paddingAngle={2}>
                {donutData.map((d, i) => <Cell key={i} fill={d.color} />)}
              </Pie>
              <Tooltip contentStyle={tooltipStyle()} formatter={(v) => fmtUGX(v)} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
            </PieChart>
          </ResponsiveContainer>
        </Card>
        <Card style={{ flex: "1 1 340px" }}>
          <SectionTitle sub="Channels ranked by cumulative revenue">Channel Breakdown</SectionTitle>
          <div style={{ display: "flex", flexDirection: "column", gap: 10, marginTop: 6 }}>
            {donutData.sort((a, b) => b.value - a.value).map((d) => (
              <div key={d.name} style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ width: 8, height: 8, borderRadius: "50%", background: d.color, flexShrink: 0 }} />
                <div style={{ flex: 1, fontSize: 13.5, color: COLORS.ink }}>{d.name}</div>
                <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 13, color: COLORS.inkSoft }}>{fmtCompact(d.value)}</div>
              </div>
            ))}
            {donutData.length === 0 && <div style={{ color: COLORS.inkFaint, fontSize: 13 }}>No data. Please add it in the 'Data Management' tab.</div>}
          </div>
        </Card>
      </div>

      {/* ---------------- ODM section ---------------- */}
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
        <div style={{ width: 10, height: 10, borderRadius: 3, background: COLORS.ochre }} />
        <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 22, fontWeight: 700, margin: 0, color: COLORS.ink }}>Custom Orders (ODM) Performance</h2>
        <span style={{ fontSize: 11, color: COLORS.inkFaint, marginLeft: 2 }}>Custom Orders (ODM)</span>
        <span style={{ fontSize: 12.5, color: COLORS.inkFaint }}>(Annual goal {fmtUGX(odmAnnualGoalUGX)})</span>
      </div>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 24 }}>
        <KpiCard icon={Briefcase} label="YTD Revenue" value={fmtCompact(odmYTD) + " UGX"} accent={COLORS.ochre} note={fmtUGX(odmYTD)} />
        <KpiCard icon={Target} label="Annual Goal Achievement" value={fmtPct(odmAchievement)} accent={odmAchievement >= 1 ? COLORS.teal : COLORS.clay} note={`Goal ${fmtCompact(odmAnnualGoalUGX)} UGX`} />
        <KpiCard icon={TrendingUp} label="Remaining to Goal" value={fmtCompact(odmRemainingGoal) + " UGX"} accent={COLORS.clay} note={odmRemainingGoal === 0 ? "Goal reached" : fmtUGX(odmRemainingGoal)} />
        <KpiCard icon={Sparkles} label="Monthly Avg. Needed" value={odmRemainingMonths > 0 ? fmtCompact(odmRequiredAvgRemaining) + " UGX" : "—"} accent={COLORS.slate} note={odmRemainingMonths > 0 ? `Based on ${odmRemainingMonths} remaining month(s)` : "12 months complete"} />
      </div>
      <Card>
        <SectionTitle sub="Bars: monthly custom-order revenue · Dashed line: annual goal divided evenly across 12 months">Custom Order Revenue Trend</SectionTitle>
        <ResponsiveContainer width="100%" height={300}>
          <ComposedChart data={odmTrendData} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} vertical={false} />
            <XAxis dataKey="period" tick={{ fontSize: 12, fill: COLORS.inkSoft }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
            <YAxis tickFormatter={fmtCompact} tick={{ fontSize: 11, fill: COLORS.inkFaint }} axisLine={false} tickLine={false} width={50} />
            <Tooltip contentStyle={tooltipStyle()} formatter={(v, name) => [fmtUGX(v), name === "amount" ? "Custom Order Revenue" : "Even Pace"]} />
            <Legend formatter={(v) => (v === "amount" ? "Custom Order Revenue" : "Even Pace (Annual Goal / 12)")} wrapperStyle={{ fontSize: 12 }} />
            <Bar dataKey="amount" fill={COLORS.ochre} radius={[4, 4, 0, 0]}>
              <LabelList dataKey="amount" content={(props) => <BarValueLabel {...props} />} />
            </Bar>
            <Line dataKey="pace" stroke={COLORS.clay} strokeWidth={2} strokeDasharray="5 4" dot={false} />
          </ComposedChart>
        </ResponsiveContainer>
        <p style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 10 }}>
          <Info size={12} style={{ verticalAlign: -1, marginRight: 4 }} />
          You can view and edit vendor-level details and inquiries still being discussed in the 'Custom Orders' tab.
        </p>
      </Card>
    </div>
  );
}

/* ---------------------------------- visits tab ---------------------------------- */
function VisitsTab({
  data, activeMonths, ytdSold, soldSeries, soldIsFromPos,
  eventEditMode, eventDraft, startEditEvents, cancelEditEvents, saveEditEvents, addEventRow, removeEventRow, updateEventRow,
}) {
  const sold = soldSeries || data.sold;
  const chartData = MONTH_LABELS.map((label, i) => ({
    period: label, visits: data.visits[i] || 0, sold: sold[i] || 0,
  }));
  const dowStats = data.dowStats || [];
  const opStats = data.operationStats || null;
  const events = eventEditMode ? eventDraft : (data.marketEvents || []);
  const eventRevenueTotal = events.reduce((a, e) => a + Number(e.revenue || 0), 0);
  // 이벤트 매출 자체가 POS Data 기준이라, 비교 분모도 POS Data 기준 매출로 통일 (Daily Report와 섞지 않음)
  const posTotalRevenue = (data.channels.pos || []).reduce((a, b) => a + Number(b || 0), 0);
  const eventConcentration = posTotalRevenue > 0 ? eventRevenueTotal / posTotalRevenue : 0;
  const avgDailyRevenue = opStats && opStats.operatingDays > 0 ? opStats.totalRevenue / opStats.operatingDays : 0;
  const originStats = data.visitorOrigin || [];
  const typeStats = data.visitorType || [];
  const highlights = buildHighlights(data);
  const topDow = dowStats.length ? [...dowStats].sort((a, b) => b.revenue - a.revenue)[0] : null;
  const bottomDow = dowStats.length ? [...dowStats].sort((a, b) => a.revenue - b.revenue)[0] : null;

  return (
    <div>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 24 }}>
        <KpiCard icon={Users} label="Total Visitors" value={fmtNum(data.visits.slice(0, activeMonths).reduce((a, b) => a + Number(b || 0), 0)) + ""} accent={COLORS.teal} />
        <KpiCard icon={ShoppingBag} label="Total Units Sold" value={fmtNum(ytdSold) + ""} accent={COLORS.clay} note={soldIsFromPos ? "Based on Product List (Details Report)" : "Based on Daily Report"} />
      </div>
      <Card style={{ marginBottom: 24 }}>
        <SectionTitle sub={`Monthly visitors and units sold (visits: Daily Report / units sold: ${soldIsFromPos ? "Product List (Details Report)" : "Daily Report"})`}>Visits & Sales Trend</SectionTitle>
        <ResponsiveContainer width="100%" height={340}>
          <ComposedChart data={chartData} margin={{ top: 24, right: 10, left: 0, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} vertical={false} />
            <XAxis dataKey="period" tick={{ fontSize: 12, fill: COLORS.inkSoft }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
            <YAxis yAxisId="left" tick={{ fontSize: 11, fill: COLORS.inkFaint }} axisLine={false} tickLine={false} />
            <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 11, fill: COLORS.inkFaint }} axisLine={false} tickLine={false} />
            <Tooltip contentStyle={tooltipStyle()} />
            <Legend wrapperStyle={{ fontSize: 12 }} formatter={(v) => ({ visits: "Visitors", sold: "Units Sold" }[v] || v)} />
            <Bar yAxisId="left" dataKey="visits" fill={COLORS.tealSoft} stroke={COLORS.teal} radius={[4, 4, 0, 0]} name="visits">
              <LabelList dataKey="visits" content={(props) => <BarValueLabel {...props} formatter={fmtNum} />} />
            </Bar>
            <Line yAxisId="right" dataKey="sold" stroke={COLORS.clay} strokeWidth={2.5} dot={{ r: 3 }} name="sold">
              <LabelList dataKey="sold" position="top" formatter={(v) => (v > 0 ? fmtNum(v) : "")} style={{ fontSize: 11, fontWeight: 600, fill: COLORS.clay }} />
            </Line>
          </ComposedChart>
        </ResponsiveContainer>
      </Card>
      <Card>
        <SectionTitle sub="Based on uploaded daily reports · cumulative performance by day of week">Visit Pattern by Day of Week</SectionTitle>
        <ResponsiveContainer width="100%" height={300}>
          <BarChart data={dowStats} margin={{ top: 24, right: 10, left: 0, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} vertical={false} />
            <XAxis dataKey="day" tick={{ fontSize: 12, fill: COLORS.inkSoft }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
            <YAxis tickFormatter={fmtCompact} tick={{ fontSize: 11, fill: COLORS.inkFaint }} axisLine={false} tickLine={false} />
            <Tooltip contentStyle={tooltipStyle()} formatter={(v, name) => [name === "revenue" ? fmtUGX(v) : v, { visits: "Visitors", revenue: "Revenue" }[name] || name]} />
            <Bar dataKey="revenue" fill={COLORS.teal} radius={[6, 6, 0, 0]} name="revenue">
              <LabelList dataKey="revenue" content={(props) => <BarValueLabel {...props} />} />
            </Bar>
          </BarChart>
        </ResponsiveContainer>
        {topDow && bottomDow && (
          <p style={{ fontSize: 12.5, color: COLORS.inkFaint, marginTop: 10 }}>
            <Info size={12} style={{ verticalAlign: -1, marginRight: 4 }} />
            {topDow.day} has the highest revenue (cumulative {fmtCompact(topDow.revenue)} UGX), and {bottomDow.day} the lowest. Useful for planning staffing or promotions by day.
          </p>
        )}
      </Card>

      {opStats && (
        <Card style={{ marginTop: 24 }}>
          <SectionTitle sub="Based on daily reports · operating-day data through the last date with results">Operational Efficiency Metrics</SectionTitle>
          <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginTop: 8 }}>
            <KpiCard icon={CalendarDays} label="Zero-Visit Day Rate" value={fmtPct(opStats.zeroVisitPct)} accent={COLORS.clay} note={`${opStats.zeroVisitDays} / ${opStats.operatingDays} days`} />
            <KpiCard icon={CalendarDays} label="Zero-Sales Day Rate" value={fmtPct(opStats.zeroSalesPct)} accent={COLORS.clay} note={`${opStats.zeroSalesDays} / ${opStats.operatingDays} days`} />
            <KpiCard icon={TrendingUp} label="Top-5-Day Revenue Concentration" value={fmtPct(opStats.top5Concentration)} accent={COLORS.ochre} note="Share of revenue from the top 5 days overall" />
            <KpiCard icon={Sparkles} label="Event Revenue Concentration" value={fmtPct(eventConcentration)} accent={COLORS.clay} note={`${events.length} event(s) · vs POS revenue`} />
          </div>
          <p style={{ fontSize: 12.5, color: COLORS.inkFaint, marginTop: 12 }}>
            <Info size={12} style={{ verticalAlign: -1, marginRight: 4 }} />
            If the zero-visit/zero-sales day rate is high, event-driven operations (markets, one-day classes) may be more efficient than staying open every day.
          </p>
        </Card>
      )}

      <Card style={{ marginTop: 24 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
          <SectionTitle sub="Based on the orange-boxed market/community event ranges in the POS Data sheet">Event Highlights</SectionTitle>
          <div style={{ display: "flex", gap: 8 }}>
            {!eventEditMode ? (
              <button onClick={startEditEvents} style={btnStyle(COLORS.teal, true)}><Pencil size={14} /> Edit</button>
            ) : (
              <>
                <button onClick={addEventRow} style={btnStyle(COLORS.ochre, true)}><Plus size={14} /> Add Event</button>
                <button onClick={saveEditEvents} style={btnStyle(COLORS.teal, true)}><Save size={14} /> Save</button>
                <button onClick={cancelEditEvents} style={btnStyle(COLORS.inkFaint, false)}><X size={14} /> Cancel</button>
              </>
            )}
          </div>
        </div>
        <div className="sdc-scroll" style={{ overflowX: "auto", marginTop: 14 }}>
          <table style={{ borderCollapse: "collapse", width: "100%", minWidth: 720, fontSize: 13 }}>
            <thead>
              <tr>
                <Th align="left">Date</Th>
                <Th align="left">Event Name</Th>
                <Th align="right">Revenue (UGX)</Th>
                <Th align="right">Units Sold</Th>
                <Th align="right">Transactions</Th>
                <Th align="right">vs Typical Day</Th>
                {eventEditMode && <Th align="center"> </Th>}
              </tr>
            </thead>
            <tbody>
              {events.map((e, idx) => (
                <tr key={e.id}>
                  <Td align="left">
                    {eventEditMode ? (
                      <input value={e.date} onChange={(ev) => updateEventRow(idx, "date", ev.target.value)} placeholder="2026-03-28"
                        style={{ width: 110, padding: "4px 6px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12.5 }} />
                    ) : e.date}
                  </Td>
                  <Td align="left">
                    {eventEditMode ? (
                      <input value={e.name} onChange={(ev) => updateEventRow(idx, "name", ev.target.value)}
                        style={{ width: 180, padding: "4px 6px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12.5 }} />
                    ) : <span style={{ fontWeight: 500 }}>{e.name}</span>}
                  </Td>
                  <Td align="right">
                    {eventEditMode ? (
                      <input type="number" value={e.revenue} onChange={(ev) => updateEventRow(idx, "revenue", ev.target.value)}
                        style={{ width: 100, padding: "4px 6px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12.5, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }} />
                    ) : <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{fmtNum(e.revenue)}</span>}
                  </Td>
                  <Td align="right">
                    {eventEditMode ? (
                      <input type="number" value={e.qty} onChange={(ev) => updateEventRow(idx, "qty", ev.target.value)}
                        style={{ width: 70, padding: "4px 6px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12.5, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }} />
                    ) : fmtNum(e.qty)}
                  </Td>
                  <Td align="right">
                    {eventEditMode ? (
                      <input type="number" value={e.invoices} onChange={(ev) => updateEventRow(idx, "invoices", ev.target.value)}
                        style={{ width: 60, padding: "4px 6px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12.5, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }} />
                    ) : fmtNum(e.invoices)}
                  </Td>
                  <Td align="right">
                    {avgDailyRevenue > 0 ? `${(e.revenue / avgDailyRevenue).toFixed(1)}x` : "—"}
                  </Td>
                  {eventEditMode && (
                    <Td align="center">
                      <button onClick={() => removeEventRow(idx)} style={{ border: "none", background: "transparent", cursor: "pointer", color: COLORS.clay }}>
                        <Trash2 size={15} />
                      </button>
                    </Td>
                  )}
                </tr>
              ))}
              {events.length === 0 && (
                <tr><td colSpan={7} style={{ padding: "16px 10px", color: COLORS.inkFaint, fontSize: 13 }}>No events registered.</td></tr>
              )}
            </tbody>
          </table>
        </div>
        <p style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 14 }}>
          <Info size={12} style={{ verticalAlign: -1, marginRight: 4 }} />
          "vs Typical Day" shows how many times the event's revenue was compared to the average daily revenue from the Daily Report ({fmtCompact(avgDailyRevenue)} UGX). (Concentration % is based on POS Data, while this multiple is based on Daily Report — the sources differ.)
        </p>
      </Card>

      <div style={{ display: "flex", alignItems: "center", gap: 8, margin: "32px 0 14px" }}>
        <div style={{ width: 10, height: 10, borderRadius: 3, background: COLORS.slate }} />
        <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 20, fontWeight: 700, margin: 0, color: COLORS.ink }}>Visitor Type Analysis</h2>
        <span style={{ fontSize: 11, color: COLORS.inkFaint, marginLeft: 2 }}>Visitor Type Analysis</span>
        <span style={{ fontSize: 12.5, color: COLORS.inkFaint }}>(based on uploaded daily report notes)</span>
      </div>

      <div style={{ display: "flex", gap: 20, flexWrap: "wrap", marginBottom: 24 }}>
        <Card style={{ flex: "1.3 1 420px" }}>
          <SectionTitle sub="Frequency of nationality/origin mentions in notes (mention count, not actual visitor count)">Visitor Origin Distribution</SectionTitle>
          <div style={{ display: "flex", flexDirection: "column", gap: 11, marginTop: 8 }}>
            {originStats.map((o) => {
              const max = originStats[0] ? originStats[0].count : 1;
              return (
                <div key={o.origin}>
                  <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 4 }}>
                    <span style={{ color: COLORS.ink }}>{o.origin}</span>
                    <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLORS.inkSoft }}>{o.count}</span>
                  </div>
                  <div style={{ height: 7, borderRadius: 4, background: COLORS.bg, overflow: "hidden" }}>
                    <div style={{ width: `${(o.count / max) * 100}%`, height: "100%", background: COLORS.slate, borderRadius: 4 }} />
                  </div>
                </div>
              );
            })}
            {originStats.length === 0 && <div style={{ color: COLORS.inkFaint, fontSize: 13 }}>No data.</div>}
          </div>
        </Card>

        <Card style={{ flex: "1 1 300px" }}>
          <SectionTitle sub="Share of 'returning/regular' vs 'new/first-time' mentions in notes">Returning vs New Visitors</SectionTitle>
          <ResponsiveContainer width="100%" height={220}>
            <PieChart>
              <Pie data={typeStats} dataKey="count" nameKey="type" innerRadius={54} outerRadius={88} paddingAngle={3}>
                {typeStats.map((d, i) => <Cell key={i} fill={d.color} />)}
              </Pie>
              <Tooltip contentStyle={tooltipStyle()} formatter={(v) => `${v} mention(s)`} />
              <Legend wrapperStyle={{ fontSize: 12 }} />
            </PieChart>
          </ResponsiveContainer>
          {typeStats.length >= 2 && typeStats[1].count > 0 && (
            <p style={{ fontSize: 12.5, color: COLORS.inkFaint, marginTop: 6, textAlign: "center" }}>
              Returning/loyal customer mentions are about {(typeStats[0].count / typeStats[1].count).toFixed(1)}x new-visitor mentions — a strong base of repeat customers.
            </p>
          )}
        </Card>
      </div>

      <Card>
        <SectionTitle sub="Recurring patterns and notable points automatically summarized from daily report notes">Notable Highlights</SectionTitle>
        <ul style={{ fontSize: 13.5, color: COLORS.ink, lineHeight: 2, paddingLeft: 20, margin: 0 }}>
          {highlights.map((h, i) => <li key={i}>{h}</li>)}
          {highlights.length === 0 && <li style={{ color: COLORS.inkFaint }}>No data to display.</li>}
        </ul>
        <p style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 14 }}>
          <Info size={12} style={{ verticalAlign: -1, marginRight: 4 }} />
          This section is a text-based summary of free-form notes from the daily reports, meant to show trends rather than exact headcounts.
        </p>
      </Card>
    </div>
  );
}

/* ---------------------------------- custom order tab ---------------------------------- */
function baseCustomerName(name) {
  return String(name || "").replace(/\s*\(.*?\)\s*/g, "").trim() || "Unspecified";
}

function CustomOrderTab({ data, editOrders, orderDraft, startEditOrders, cancelEditOrders, saveEditOrders, addOrderRow, removeOrderRow, updateOrderRow, addProductLineToOrder, reloadOrdersFromDefaults, editLeads, leadsDraft, startEditLeads, cancelEditLeads, saveEditLeads, addLead, removeLead, updateLead }) {
  const [confirmReload, setConfirmReload] = useState(false);
  const orders = data.customOrders || [];
  const shown = editOrders ? orderDraft : orders;

  const totalFromChannel = data.channels.custom.reduce((a, b) => a + Number(b || 0), 0);
  // 한 주문을 상품별로 여러 줄에 나눠 입력해도 실제 주문 건수는 groupId 기준으로 정확히 셈
  const orderCount = new Set(orders.map((o) => o.groupId ?? o.id)).size;
  const sumFromOrders = orders.reduce((a, o) => a + Number(o.amount || 0), 0);
  const avgOrder = orderCount > 0 ? sumFromOrders / orderCount : 0;

  const byCustomer = {};
  orders.forEach((o) => {
    const key = baseCustomerName(o.customer);
    if (!byCustomer[key]) byCustomer[key] = { amount: 0, groupIds: new Set(), months: new Set() };
    byCustomer[key].amount += Number(o.amount || 0);
    byCustomer[key].groupIds.add(o.groupId ?? o.id);
    byCustomer[key].months.add(o.month);
  });
  const ranking = Object.entries(byCustomer)
    .map(([name, v]) => ({ name, amount: v.amount, orderCount: v.groupIds.size, monthCount: v.months.size, isRepeat: v.groupIds.size >= 2 }))
    .sort((a, b) => b.amount - a.amount);
  const topCustomer = ranking[0];
  const maxRank = ranking.length ? ranking[0].amount : 1;
  const repeatCustomers = ranking.filter((c) => c.isRepeat);
  const repeatShare = ranking.length > 0 ? repeatCustomers.reduce((a, c) => a + c.amount, 0) / ranking.reduce((a, c) => a + c.amount, 0) : 0;

  const byCategory = {};
  orders.forEach((o) => {
    const cat = o.category || "Unspecified";
    byCategory[cat] = (byCategory[cat] || 0) + Number(o.amount || 0);
  });
  const categoryBreakdown = Object.entries(byCategory)
    .map(([cat, amount]) => ({ cat, amount, count: orders.filter((o) => (o.category || "Unspecified") === cat).length }))
    .sort((a, b) => b.amount - a.amount);
  const categoryTotal = categoryBreakdown.reduce((a, c) => a + c.amount, 0);

  const byProduct = {};
  orders.forEach((o) => {
    if (!o.product) return;
    const key = o.product.trim();
    if (!key) return;
    if (!byProduct[key]) byProduct[key] = { product: key, qty: 0, count: 0 };
    byProduct[key].qty += Number(o.qty || 0);
    byProduct[key].count += 1;
  });
  const productBreakdown = Object.values(byProduct).sort((a, b) => b.qty - a.qty);
  const maxProductQty = productBreakdown.length ? Math.max(...productBreakdown.map((p) => p.qty)) : 1;
  const taggedOrderCount = orders.filter((o) => o.product).length;

  const trendData = MONTH_LABELS.map((label, i) => ({
    period: label,
    custom: data.channels.custom[i] || 0,
    total: monthTotal(data, i),
  }));

  const reconciliation = MONTH_LABELS
    .map((label, i) => {
      const recorded = orders.filter((o) => o.month === i).reduce((a, o) => a + Number(o.amount || 0), 0);
      const channelTotal = data.channels.custom[i] || 0;
      return { label, recorded, channelTotal, diff: channelTotal - recorded };
    })
    .filter((r) => r.channelTotal > 0 || r.recorded > 0);
  const mismatches = reconciliation.filter((r) => Math.abs(r.diff) > 1);

  return (
    <div>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 24 }}>
        <KpiCard icon={Briefcase} label="Total Custom Order Revenue" value={fmtCompact(totalFromChannel) + " UGX"} accent={COLORS.ochre} note={fmtUGX(totalFromChannel)} />
        <KpiCard icon={TrendingUp} label="Order Count" value={fmtNum(orderCount) + ""} accent={COLORS.teal} note={`Avg ${fmtCompact(avgOrder)} UGX / order`} />
        <KpiCard icon={Target} label="Top Customer" value={topCustomer ? topCustomer.name : "—"} accent={COLORS.clay} note={topCustomer ? fmtUGX(topCustomer.amount) : ""} />
        <KpiCard icon={CalendarDays} label="Inquiries Being Discussed" value={fmtNum((data.customLeads || []).length) + ""} accent={COLORS.slate} note="Amount not yet confirmed" />
        <KpiCard icon={Repeat} label="Repeat Customer Revenue Share" value={ranking.length > 0 ? fmtPct(repeatShare) : "—"} accent={COLORS.teal} note={`${repeatCustomers.length} / ${ranking.length} are repeat`} />
      </div>

      <Card style={{ marginBottom: 24 }}>
        <SectionTitle sub="Custom order revenue alongside its share of total revenue">Monthly Custom Order Trend</SectionTitle>
        <ResponsiveContainer width="100%" height={320}>
          <ComposedChart data={trendData} margin={{ top: 24, right: 10, left: 0, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} vertical={false} />
            <XAxis dataKey="period" tick={{ fontSize: 12, fill: COLORS.inkSoft }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
            <YAxis tickFormatter={fmtCompact} tick={{ fontSize: 11, fill: COLORS.inkFaint }} axisLine={false} tickLine={false} width={50} />
            <Tooltip contentStyle={tooltipStyle()} formatter={(v, name) => [fmtUGX(v), name === "custom" ? "Custom Orders" : "Total Revenue"]} />
            <Legend formatter={(v) => (v === "custom" ? "Custom Orders" : "Total Revenue")} wrapperStyle={{ fontSize: 12 }} />
            <Bar dataKey="custom" fill={COLORS.ochre} radius={[4, 4, 0, 0]}>
              <LabelList dataKey="custom" content={(props) => <BarValueLabel {...props} />} />
            </Bar>
            <Line dataKey="total" stroke={COLORS.slate} strokeWidth={2} strokeDasharray="4 3" dot={false} />
          </ComposedChart>
        </ResponsiveContainer>
      </Card>

      <div style={{ display: "flex", gap: 20, flexWrap: "wrap", marginBottom: 24 }}>
        <Card style={{ flex: "1 1 420px" }}>
          <SectionTitle sub="Cumulative order amount by customer (color variants are merged under the same customer) · badge shows reorder count">Customer Ranking</SectionTitle>
          {ranking.length > 0 && (
            <p style={{ fontSize: 12.5, color: COLORS.inkFaint, margin: "4px 0 12px" }}>
              {repeatCustomers.length} repeat customer(s) account for <b style={{ color: COLORS.teal }}>{fmtPct(repeatShare)}</b> of total custom order revenue.
            </p>
          )}
          <div style={{ display: "flex", flexDirection: "column", gap: 12, marginTop: 8 }}>
            {ranking.map((c) => (
              <div key={c.name}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 4 }}>
                  <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <span style={{ color: COLORS.ink, fontWeight: 500 }}>{c.name}</span>
                    {c.isRepeat ? (
                      <span style={{ fontSize: 11, fontWeight: 700, color: COLORS.teal, background: COLORS.tealSoft, padding: "1px 8px", borderRadius: 10 }}>
                        Reordered {c.orderCount}x
                      </span>
                    ) : (
                      <span style={{ fontSize: 11, color: COLORS.inkFaint, background: COLORS.bg, padding: "1px 8px", borderRadius: 10 }}>One-time</span>
                    )}
                  </span>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLORS.inkSoft }}>{fmtCompact(c.amount)}</span>
                </div>
                <div style={{ height: 7, borderRadius: 4, background: COLORS.bg, overflow: "hidden" }}>
                  <div style={{ width: `${(c.amount / maxRank) * 100}%`, height: "100%", background: c.isRepeat ? COLORS.teal : COLORS.ochre, borderRadius: 4 }} />
                </div>
              </div>
            ))}
            {ranking.length === 0 && <div style={{ color: COLORS.inkFaint, fontSize: 13 }}>No orders registered.</div>}
          </div>
        </Card>
        <Card style={{ flex: "1 1 280px" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
            <SectionTitle sub="Inquiries with amounts not yet confirmed">Customers Being Discussed</SectionTitle>
            <div style={{ display: "flex", gap: 6 }}>
              {!editLeads ? (
                <button onClick={startEditLeads} style={btnStyle(COLORS.teal, true)}><Pencil size={12} /> Edit</button>
              ) : (
                <>
                  <button onClick={addLead} style={btnStyle(COLORS.ochre, true)}><Plus size={12} /> Add</button>
                  <button onClick={saveEditLeads} style={btnStyle(COLORS.teal, true)}><Save size={12} /> Save</button>
                  <button onClick={cancelEditLeads} style={btnStyle(COLORS.inkFaint, false)}><X size={12} /> Cancel</button>
                </>
              )}
            </div>
          </div>
          {!editLeads ? (
            <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginTop: 10 }}>
              {(data.customLeads || []).map((l) => (
                <span key={l} style={{
                  padding: "6px 12px", borderRadius: 20, background: COLORS.tealSoft, color: COLORS.teal,
                  fontSize: 12.5, fontWeight: 600,
                }}>{l}</span>
              ))}
              {(!data.customLeads || data.customLeads.length === 0) && <div style={{ color: COLORS.inkFaint, fontSize: 13 }}>None</div>}
            </div>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 10 }}>
              {leadsDraft.map((l, idx) => (
                <div key={idx} style={{ display: "flex", gap: 6, alignItems: "center" }}>
                  <input
                    value={l} onChange={(e) => updateLead(idx, e.target.value)} placeholder="Customer name"
                    style={{ flex: 1, padding: "6px 10px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 13 }}
                  />
                  <button onClick={() => removeLead(idx)} style={{ border: "none", background: "transparent", cursor: "pointer", color: COLORS.clay }}>
                    <Trash2 size={15} />
                  </button>
                </div>
              ))}
              {leadsDraft.length === 0 && <div style={{ color: COLORS.inkFaint, fontSize: 13 }}>No customers registered. Use the 'Add' button.</div>}
            </div>
          )}
          {mismatches.length > 0 ? (
            <div style={{ marginTop: 18, padding: 12, borderRadius: 8, background: COLORS.claySoft, fontSize: 12.5, color: COLORS.ink }}>
              <b>Needs review:</b> {mismatches.map((m) => `${m.label} (diff ${fmtCompact(m.diff)})`).join(", ")} — the channel total doesn't match the individual order records. Please add any missing orders in the table below.
            </div>
          ) : (
            <div style={{ marginTop: 18, padding: 12, borderRadius: 8, background: COLORS.tealSoft, fontSize: 12.5, color: COLORS.teal }}>
              Every month's order records match the channel revenue.
            </div>
          )}
        </Card>
      </div>

      <Card style={{ marginBottom: 24 }}>
        <SectionTitle sub="Based on the customer category recorded in each order">Customer Type Analysis</SectionTitle>
        <div style={{ display: "flex", gap: 20, flexWrap: "wrap", marginTop: 8 }}>
          <div style={{ flex: "1 1 260px" }}>
            <ResponsiveContainer width="100%" height={240}>
              <PieChart>
                <Pie data={categoryBreakdown} dataKey="amount" nameKey="cat" innerRadius={52} outerRadius={84} paddingAngle={2}>
                  {categoryBreakdown.map((c, i) => <Cell key={i} fill={CUSTOMER_CATEGORY_COLORS[c.cat] || COLORS.inkFaint} />)}
                </Pie>
                <Tooltip contentStyle={tooltipStyle()} formatter={(v) => fmtUGX(v)} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
              </PieChart>
            </ResponsiveContainer>
          </div>
          <div style={{ flex: "1 1 260px", display: "flex", flexDirection: "column", gap: 10, justifyContent: "center" }}>
            {categoryBreakdown.map((c) => (
              <div key={c.cat}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 4 }}>
                  <span style={{ color: COLORS.ink }}>{c.cat}</span>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLORS.inkSoft }}>
                    {fmtCompact(c.amount)} · {c.count} order(s) · {fmtPct(categoryTotal > 0 ? c.amount / categoryTotal : 0)}
                  </span>
                </div>
                <div style={{ height: 7, borderRadius: 4, background: COLORS.bg, overflow: "hidden" }}>
                  <div style={{ width: `${categoryTotal > 0 ? (c.amount / categoryTotal) * 100 : 0}%`, height: "100%", background: CUSTOMER_CATEGORY_COLORS[c.cat] || COLORS.inkFaint, borderRadius: 4 }} />
                </div>
              </div>
            ))}
          </div>
        </div>
        {byCategory["Unspecified"] > 0 && (
          <p style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 12 }}>
            <Info size={12} style={{ verticalAlign: -1, marginRight: 4 }} />
            For orders with no category in the 'Order History' table below, assign one of Hotel/Restaurant, Government Office, NGO/Institution, Individual, or Other for more accurate analysis.
          </p>
        )}
      </Card>

      <Card style={{ marginBottom: 24 }}>
        <SectionTitle sub="Based on orders with a 'Product' entered in the order table · by quantity/count (revenue is excluded since a multi-item order's amount isn't always split per item)">Custom Order Product Analysis</SectionTitle>
        {productBreakdown.length > 0 ? (
          <div style={{ display: "flex", flexDirection: "column", gap: 12, marginTop: 8 }}>
            {productBreakdown.map((p) => (
              <div key={p.product}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 4 }}>
                  <span style={{ color: COLORS.ink, fontWeight: 500 }}>{p.product}</span>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLORS.inkSoft }}>
                    {fmtNum(p.qty)} · {p.count} order(s)
                  </span>
                </div>
                <div style={{ height: 7, borderRadius: 4, background: COLORS.bg, overflow: "hidden" }}>
                  <div style={{ width: `${(p.qty / maxProductQty) * 100}%`, height: "100%", background: COLORS.ochre, borderRadius: 4 }} />
                </div>
              </div>
            ))}
          </div>
        ) : (
          <p style={{ color: COLORS.inkFaint, fontSize: 13, marginTop: 8 }}>
            No orders have a product entered yet. In the 'Order History' table below, switch to edit mode and enter a product name and quantity for each order. (e.g., Korea Embassy → product "eco bag", qty 100)
          </p>
        )}
        {taggedOrderCount > 0 && taggedOrderCount < orderCount && (
          <p style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 14 }}>
            <Info size={12} style={{ verticalAlign: -1, marginRight: 4 }} />
            Only {taggedOrderCount} of {orderCount} orders have a product entered. Filling in the rest will give more accurate analysis.
          </p>
        )}
      </Card>

      <Card>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
          <SectionTitle sub="Jan-Jun entries were compiled from the original Excel cell notes. When a new order comes in, add it here along with category, product and quantity. If an order has multiple products, use the row's + button to add another product line to the same order (it won't be double-counted in the order count).">
            Order History
          </SectionTitle>
          <div style={{ display: "flex", gap: 8 }}>
            {!editOrders ? (
              <>
                <button onClick={startEditOrders} style={btnStyle(COLORS.teal, true)}><Pencil size={14} /> Edit</button>
                {!confirmReload ? (
                  <button onClick={() => setConfirmReload(true)} style={btnStyle(COLORS.inkFaint, false)}>
                    <RotateCcw size={14} /> Reload Latest Defaults
                  </button>
                ) : (
                  <>
                    <span style={{ fontSize: 12.5, color: COLORS.clay, alignSelf: "center", fontWeight: 600 }}>Really overwrite?</span>
                    <button onClick={() => { reloadOrdersFromDefaults(); setConfirmReload(false); }} style={btnStyle(COLORS.clay, true)}>
                      <RotateCcw size={14} /> Yes, reload
                    </button>
                    <button onClick={() => setConfirmReload(false)} style={btnStyle(COLORS.inkFaint, false)}>
                      <X size={14} /> Cancel
                    </button>
                  </>
                )}
              </>
            ) : (
              <>
                <button onClick={addOrderRow} style={btnStyle(COLORS.ochre, true)}><Plus size={14} /> Add Order</button>
                <button onClick={saveEditOrders} style={btnStyle(COLORS.teal, true)}><Save size={14} /> Save</button>
                <button onClick={cancelEditOrders} style={btnStyle(COLORS.inkFaint, false)}><X size={14} /> Cancel</button>
              </>
            )}
          </div>
        </div>
        <div className="sdc-scroll" style={{ overflowX: "auto", marginTop: 14 }}>
          <table style={{ borderCollapse: "collapse", width: "100%", minWidth: 640, fontSize: 13 }}>
            <thead>
              <tr>
                <Th align="left">Month</Th>
                <Th align="left">Customer</Th>
                <Th align="left">Category</Th>
                <Th align="left">Product</Th>
                <Th align="right">Qty</Th>
                <Th align="right">Amount (UGX)</Th>
                <Th align="left">Note</Th>
                {editOrders && <Th align="center"> </Th>}
              </tr>
            </thead>
            <tbody>
              {shown.map((o, idx) => (
                <tr key={o.id}>
                  <Td align="left">
                    {editOrders ? (
                      <select value={o.month} onChange={(e) => updateOrderRow(idx, "month", e.target.value)}
                        style={{ padding: "4px 6px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12.5 }}>
                        {MONTH_LABELS.map((m, i) => <option key={m} value={i}>{m}</option>)}
                      </select>
                    ) : MONTH_LABELS[o.month]}
                  </Td>
                  <Td align="left">
                    {editOrders ? (
                      <input value={o.customer} onChange={(e) => updateOrderRow(idx, "customer", e.target.value)}
                        style={{ width: 150, padding: "4px 6px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12.5 }} />
                    ) : <span style={{ fontWeight: 500 }}>{o.customer}</span>}
                  </Td>
                  <Td align="left">
                    {editOrders ? (
                      <select value={o.category || ""} onChange={(e) => updateOrderRow(idx, "category", e.target.value || null)}
                        style={{ padding: "4px 6px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12.5 }}>
                        <option value="">Unspecified</option>
                        {CUSTOMER_CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
                      </select>
                    ) : (
                      o.category ? (
                        <span style={{
                          padding: "3px 9px", borderRadius: 14, fontSize: 11.5, fontWeight: 600,
                          background: (CUSTOMER_CATEGORY_COLORS[o.category] || COLORS.inkFaint) + "22",
                          color: CUSTOMER_CATEGORY_COLORS[o.category] || COLORS.inkFaint,
                        }}>{o.category}</span>
                      ) : <span style={{ color: COLORS.inkFaint, fontSize: 12 }}>Unspecified</span>
                    )}
                  </Td>
                  <Td align="left">
                    {editOrders ? (
                      <input value={o.product || ""} onChange={(e) => updateOrderRow(idx, "product", e.target.value)}
                        style={{ width: 130, padding: "4px 6px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12.5 }} />
                    ) : (o.product || <span style={{ color: COLORS.inkFaint, fontSize: 12 }}>—</span>)}
                  </Td>
                  <Td align="right">
                    {editOrders ? (
                      <input type="number" value={o.qty === null || o.qty === undefined ? "" : o.qty} onChange={(e) => updateOrderRow(idx, "qty", e.target.value)}
                        style={{ width: 70, padding: "4px 6px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12.5, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }} />
                    ) : (o.qty != null ? fmtNum(o.qty) : <span style={{ color: COLORS.inkFaint, fontSize: 12 }}>—</span>)}
                  </Td>
                  <Td align="right">
                    {editOrders ? (
                      <input type="number" value={o.amount} onChange={(e) => updateOrderRow(idx, "amount", e.target.value)}
                        style={{ width: 110, padding: "4px 6px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12.5, textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }} />
                    ) : <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{fmtNum(o.amount)}</span>}
                  </Td>
                  <Td align="left">
                    {editOrders ? (
                      <input value={o.note} onChange={(e) => updateOrderRow(idx, "note", e.target.value)}
                        style={{ width: 200, padding: "4px 6px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12.5 }} />
                    ) : <span style={{ color: COLORS.inkFaint }}>{o.note}</span>}
                  </Td>
                  {editOrders && (
                    <Td align="center">
                      <div style={{ display: "flex", gap: 6, justifyContent: "center" }}>
                        <button onClick={() => addProductLineToOrder(idx)} title="Add a product line to this order" style={{ border: "none", background: "transparent", cursor: "pointer", color: COLORS.teal }}>
                          <Plus size={15} />
                        </button>
                        <button onClick={() => removeOrderRow(idx)} title="Delete this row" style={{ border: "none", background: "transparent", cursor: "pointer", color: COLORS.clay }}>
                          <Trash2 size={15} />
                        </button>
                      </div>
                    </Td>
                  )}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

/* ---------------------------------- vendor (producer) analysis tab ---------------------------------- */
function VendorTab({ data, toggleConsignmentVendor }) {
  const [month, setMonth] = useState(null);
  const baseVendorStats = data.vendorStats || [];
  const vendorMonthly = data.vendorMonthly || {};
  const vendorMonthlyDetail = data.vendorMonthlyDetail || {};

  const vendorStats = useMemo(() => {
    if (month === null) return baseVendorStats;
    return baseVendorStats
      .map((v) => {
        const d = (vendorMonthlyDetail[v.vendor] || [])[month];
        return d ? { vendor: v.vendor, qty: d.qty, gross: d.gross, net: d.net, discount: d.discount, count: d.count } : null;
      })
      .filter((v) => v && (v.net > 0 || v.qty > 0))
      .sort((a, b) => b.net - a.net);
  }, [month, baseVendorStats, vendorMonthlyDetail]);

  const totalNet = vendorStats.reduce((a, v) => a + v.net, 0);
  const totalQty = vendorStats.reduce((a, v) => a + v.qty, 0);
  const sdcStat = vendorStats.find((v) => v.vendor === "SDC");
  const sdcShare = totalNet > 0 && sdcStat ? sdcStat.net / totalNet : 0;
  const consignmentShare = totalNet > 0 ? 1 - sdcShare : 0;
  const topVendor = vendorStats[0];

  const donutData = vendorStats.map((v, i) => ({ name: v.vendor, value: v.net, color: VENDOR_COLORS[i % VENDOR_COLORS.length] }));

  const trendData = MONTH_LABELS.map((label, i) => {
    const row = { period: label };
    baseVendorStats.forEach((v) => { row[v.vendor] = (vendorMonthlyDetail[v.vendor] || [])[i]?.net || 0; });
    row.__labelTotal = baseVendorStats.reduce((s, v) => s + row[v.vendor], 0) > 0 ? 1 : 0;
    return row;
  });

  const detailRows = vendorStats.map((v) => ({
    ...v,
    share: totalNet > 0 ? v.net / totalNet : 0,
    avgPrice: v.qty > 0 ? v.net / v.qty : 0,
    discountRate: v.gross > 0 ? v.discount / v.gross : 0,
  }));

  const hasMonthlyDetail = Object.keys(vendorMonthlyDetail).length > 0;

  return (
    <div>
      {baseVendorStats.length === 0 ? (
        <Card>
          <p style={{ color: COLORS.inkFaint, fontSize: 13.5 }}>
            No vendor (category) data yet. Please upload a sales data Excel with a 'Category' column in the POS Data sheet, in the 'Data Management' tab.
          </p>
        </Card>
      ) : (
        <>
          <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 14 }}>
            <MonthPicker month={month} setMonth={setMonth} />
          </div>

          {month !== null && !hasMonthlyDetail ? (
            <Card>
              <p style={{ color: COLORS.inkFaint, fontSize: 13.5 }}>
                No monthly vendor detail yet. Uploading the sales data Excel once more in the 'Data Management' tab will enable monthly viewing.
              </p>
            </Card>
          ) : (
            <>
              <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 24 }}>
                <KpiCard icon={Factory} label="Number of Vendors" value={fmtNum(vendorStats.length) + ""} accent={COLORS.teal} />
                <KpiCard icon={Target} label="SDC In-House Share" value={fmtPct(sdcShare)} accent={COLORS.teal} note={sdcStat ? fmtUGX(sdcStat.net) : ""} />
                <KpiCard icon={Briefcase} label="Consignment Share" value={fmtPct(consignmentShare)} accent={COLORS.ochre} note="Sum of all vendors other than SDC" />
                <KpiCard icon={TrendingUp} label="Top-Revenue Vendor" value={topVendor ? topVendor.vendor : "—"} accent={COLORS.clay} note={topVendor ? fmtUGX(topVendor.net) : ""} />
              </div>

              <div style={{
                padding: 14, borderRadius: 10, background: COLORS.tealSoft, fontSize: 12.5, color: COLORS.ink, marginBottom: 24,
              }}>
                <Info size={13} style={{ verticalAlign: -1, marginRight: 5 }} />
                <b>Apoyo</b> is a consignment arrangement settled after the fact, based on what sells. Gift shop revenue (POS) includes Apoyo's sales as-is, but part of it is later settled to Apoyo, so SDC's actual take may differ. Other vendors' products are paid for upfront by SDC, so revenue belongs fully to SDC the moment they sell.
                The <b>"Consignment (settled after sale)"</b> button in the table below is just a reference tag marking which vendors use this structure (it doesn't affect the gift shop revenue calculation).
              </div>

              <div style={{ display: "flex", gap: 20, flexWrap: "wrap", marginBottom: 24 }}>
                <Card style={{ flex: "1 1 340px" }}>
                  <SectionTitle sub={month === null ? "Cumulative revenue share by vendor" : `${MONTH_LABELS[month]} revenue share`}>Revenue Share</SectionTitle>
                  <ResponsiveContainer width="100%" height={280}>
                    <PieChart>
                      <Pie data={donutData} dataKey="value" nameKey="name" innerRadius={62} outerRadius={98} paddingAngle={2}>
                        {donutData.map((d, i) => <Cell key={i} fill={d.color} />)}
                      </Pie>
                      <Tooltip contentStyle={tooltipStyle()} formatter={(v) => fmtUGX(v)} />
                      <Legend wrapperStyle={{ fontSize: 12 }} />
                    </PieChart>
                  </ResponsiveContainer>
                </Card>
                <Card style={{ flex: "1 1 340px" }}>
                  <SectionTitle sub="Share by units sold">Quantity Share</SectionTitle>
                  <div style={{ display: "flex", flexDirection: "column", gap: 10, marginTop: 6 }}>
                    {vendorStats.map((v, i) => (
                      <div key={v.vendor}>
                        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 4 }}>
                          <span style={{ color: COLORS.ink }}>{v.vendor}</span>
                          <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLORS.inkSoft }}>{fmtNum(v.qty)} · {fmtPct(totalQty > 0 ? v.qty / totalQty : 0)}</span>
                        </div>
                        <div style={{ height: 7, borderRadius: 4, background: COLORS.bg, overflow: "hidden" }}>
                          <div style={{ width: `${totalQty > 0 ? (v.qty / totalQty) * 100 : 0}%`, height: "100%", background: VENDOR_COLORS[i % VENDOR_COLORS.length], borderRadius: 4 }} />
                        </div>
                      </div>
                    ))}
                  </div>
                </Card>
              </div>

              <Card style={{ marginBottom: 24 }}>
                <SectionTitle sub="Monthly revenue trend by vendor (based on POS transactions · always shows the full year)">Vendor Revenue Trend</SectionTitle>
                <ResponsiveContainer width="100%" height={320}>
                  <ComposedChart data={trendData} margin={{ top: 24, right: 10, left: 0, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} vertical={false} />
                    <XAxis dataKey="period" tick={{ fontSize: 12, fill: COLORS.inkSoft }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
                    <YAxis tickFormatter={fmtCompact} tick={{ fontSize: 11, fill: COLORS.inkFaint }} axisLine={false} tickLine={false} width={50} />
                    <Tooltip contentStyle={tooltipStyle()} formatter={(v) => fmtUGX(v)} />
                    <Legend wrapperStyle={{ fontSize: 12 }} />
                    {baseVendorStats.map((v, i) => (
                      <Bar key={v.vendor} dataKey={v.vendor} stackId="vendor" fill={VENDOR_COLORS[i % VENDOR_COLORS.length]} radius={0} />
                    ))}
                    <Bar dataKey="__labelTotal" stackId="vendor" fill="transparent" radius={0} isAnimationActive={false}>
                      <LabelList dataKey="__labelTotal" content={(props) => <StackTotalLabel {...props} rows={trendData} keys={baseVendorStats.map((vv) => vv.vendor)} anchorKey="__labelTotal" />} />
                    </Bar>
                  </ComposedChart>
                </ResponsiveContainer>
              </Card>

              <Card>
                <SectionTitle sub={month === null ? "Revenue, quantity, avg. price, and discount rate by vendor (cumulative)" : `${MONTH_LABELS[month]} revenue, quantity, avg. price, and discount rate by vendor`}>Vendor Detail</SectionTitle>
                <div className="sdc-scroll" style={{ overflowX: "auto", marginTop: 10 }}>
                  <table style={{ borderCollapse: "collapse", width: "100%", minWidth: 640, fontSize: 13 }}>
                    <thead>
                      <tr>
                        <Th align="left">Vendor</Th>
                        <Th align="right">Revenue</Th>
                        <Th align="right">Share</Th>
                        <Th align="right">Units Sold</Th>
                        <Th align="right">Avg. Price</Th>
                        <Th align="right">Discount Rate</Th>
                        <Th align="right">Transactions</Th>
                        <Th align="center">Consignment (settled after sale)</Th>
                      </tr>
                    </thead>
                    <tbody>
                      {detailRows.map((v) => {
                        const isConsignment = (data.consignmentVendors || []).includes(v.vendor);
                        return (
                        <tr key={v.vendor}>
                          <Td align="left" bold>{v.vendor}</Td>
                          <Td align="right">{fmtNum(v.net)}</Td>
                          <Td align="right">{fmtPct(v.share)}</Td>
                          <Td align="right">{fmtNum(v.qty)}</Td>
                          <Td align="right">{fmtNum(v.avgPrice)}</Td>
                          <Td align="right">{fmtPct(v.discountRate)}</Td>
                          <Td align="right">{fmtNum(v.count)}</Td>
                          <Td align="center">
                            <button
                              onClick={() => toggleConsignmentVendor(v.vendor)}
                              style={{
                                padding: "4px 12px", borderRadius: 14, fontSize: 11.5, fontWeight: 600, cursor: "pointer",
                                border: isConsignment ? "none" : `1px solid ${COLORS.border}`,
                                background: isConsignment ? COLORS.clay : "transparent",
                                color: isConsignment ? "#fff" : COLORS.inkFaint,
                              }}
                            >
                              {isConsignment ? "Consignment (post-sale)" : "Paid upfront"}
                            </button>
                          </Td>
                        </tr>
                        );
                      })}
                </tbody>
              </table>
            </div>
            <p style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 14 }}>
              <Info size={12} style={{ verticalAlign: -1, marginRight: 4 }} />
              Without production quantity/inventory data, a "sell-through rate (sold vs. produced)" can't be calculated, so vendor trends are instead compared using revenue/quantity share and avg. price/discount rate.
              A vendor with a higher average price tends toward low-volume, high-price sales; a higher discount rate suggests more room for promotions or negotiation.
            </p>
          </Card>
            </>
          )}
        </>
      )}
    </div>
  );
}

/* ---------------------------------- products tab ---------------------------------- */
function ProductsTab({ data, groupFormOpen, setGroupFormOpen, groupNameInput, setGroupNameInput, groupMembersInput, setGroupMembersInput, addProductGroup, removeProductGroup }) {
  const [month, setMonth] = useState(null);
  const [perfVendorFilter, setPerfVendorFilter] = useState("all");
  const [bestsellerVendorFilter, setBestsellerVendorFilter] = useState("all");
  const topProducts = data.topProducts || [];
  const allProducts = data.allProducts || [];
  const paymentStats = data.paymentStats || [];
  const perf = data.productPerformance || null;
  const productGroups = data.productGroups || [];
  const itemVendorMap = data.itemVendorMap || {};

  const vendorOptions = useMemo(() => {
    const set = new Set(Object.values(itemVendorMap));
    return [...set].sort();
  }, [itemVendorMap]);

  const filteredPerf = useMemo(() => {
    if (!perf) return { totalCount: 0, zeroOrOneCount: 0, lessThan5Count: 0, underperformers: [] };
    const all = perf.allSellable || [];
    if (perfVendorFilter === "all" || all.length === 0) {
      return { totalCount: perf.totalCount, zeroOrOneCount: perf.zeroOrOneCount, lessThan5Count: perf.lessThan5Count, underperformers: perf.underperformers };
    }
    const matches = all.filter((p) => itemVendorMap[normalizeProductNameKey(p.name)] === perfVendorFilter);
    const zeroOrOne = matches.filter((p) => p.qty <= 1);
    const lessThan5 = matches.filter((p) => p.qty < 5);
    return {
      totalCount: matches.length,
      zeroOrOneCount: zeroOrOne.length,
      lessThan5Count: lessThan5.length,
      underperformers: [...lessThan5].sort((a, b) => a.qty - b.qty).slice(0, 15).map((p) => `${p.name}(${p.qty})`),
    };
  }, [perf, perfVendorFilter, itemVendorMap]);

  const displayProducts = useMemo(() => {
    let base;
    if (month === null) {
      // '전체' 보기: topProducts(스프레드시트 자체 합계, 가장 정확)는 그대로 신뢰하고,
      // 순위 밖이라 topProducts에 없는 변형 상품만 allProducts에서 보충해서 그룹핑 시 누락되지 않게 함
      if (allProducts.length > 0) {
        const topNames = new Set(topProducts.map((p) => p.name));
        const supplement = allProducts
          .filter((p) => !topNames.has(p.name))
          .map((p) => ({
            name: p.name,
            qty: (p.qty || []).reduce((a, b) => a + (Number(b) || 0), 0),
            amount: (p.amount || []).reduce((a, b) => a + (Number(b) || 0), 0),
          }))
          .filter((p) => p.amount > 0);
        base = [...topProducts, ...supplement];
      } else {
        base = topProducts;
      }
    } else if (allProducts.length === 0) {
      base = [];
    } else {
      base = allProducts
        .map((p) => ({ name: p.name, qty: p.qty[month] || 0, amount: p.amount[month] || 0 }))
        .filter((p) => p.amount > 0);
    }
    if (bestsellerVendorFilter !== "all") {
      base = base.filter((p) => itemVendorMap[normalizeProductNameKey(p.name)] === bestsellerVendorFilter);
    }
    const grouped = applyProductGrouping(base, productGroups);
    return grouped.sort((a, b) => b.amount - a.amount).slice(0, 10);
  }, [month, topProducts, allProducts, productGroups, bestsellerVendorFilter, itemVendorMap]);

  const maxAmt = displayProducts.length ? Math.max(...displayProducts.map((p) => p.amount)) : 1;
  const totalPayment = paymentStats.reduce((a, p) => a + p.amount, 0);
  const cashCardShare = totalPayment > 0
    ? paymentStats.filter((p) => /cash|card|현금|카드/i.test(p.method)).reduce((a, p) => a + p.amount, 0) / totalPayment
    : 0;
  return (
    <div>
      <div style={{ display: "flex", gap: 20, flexWrap: "wrap", marginBottom: 24 }}>
        <Card style={{ flex: "1.4 1 420px" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 4 }}>
            <SectionTitle sub="Top products by sales amount">Best-Selling Products</SectionTitle>
            <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
              {vendorOptions.length > 0 && (
                <select
                  value={bestsellerVendorFilter} onChange={(e) => setBestsellerVendorFilter(e.target.value)}
                  style={{ padding: "6px 10px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12.5, color: COLORS.ink, background: "#fff" }}
                >
                  <option value="all">All (every vendor)</option>
                  {vendorOptions.map((v) => <option key={v} value={v}>{v} only</option>)}
                </select>
              )}
              <MonthPicker month={month} setMonth={setMonth} />
            </div>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 12, marginTop: 8 }}>
            {displayProducts.map((p, i) => (
              <div key={p.name}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, marginBottom: 4 }}>
                  <span style={{ color: COLORS.ink, fontWeight: 500 }}>{i + 1}. {p.name}</span>
                  <span style={{ fontFamily: "'IBM Plex Mono', monospace", color: COLORS.inkSoft }}>{fmtCompact(p.amount)} · {p.qty}</span>
                </div>
                <div style={{ height: 7, borderRadius: 4, background: COLORS.bg, overflow: "hidden" }}>
                  <div style={{ width: `${(p.amount / maxAmt) * 100}%`, height: "100%", background: COLORS.teal, borderRadius: 4 }} />
                </div>
              </div>
            ))}
            {displayProducts.length === 0 && (
              <div style={{ color: COLORS.inkFaint, fontSize: 13 }}>
                {month === null ? "No data." : "No sales data for that month. (Monthly product data may need one more Excel re-upload.)"}
              </div>
            )}
          </div>
        </Card>
        <Card style={{ flex: "1 1 320px" }}>
          <SectionTitle sub="Revenue share by payment method (cumulative)">Payment Method Mix</SectionTitle>
          <ResponsiveContainer width="100%" height={260}>
            <PieChart>
              <Pie data={paymentStats} dataKey="amount" nameKey="method" innerRadius={58} outerRadius={92} paddingAngle={2}>
                {paymentStats.map((d, i) => <Cell key={i} fill={PAY_COLORS[i % PAY_COLORS.length]} />)}
              </Pie>
              <Tooltip contentStyle={tooltipStyle()} formatter={(v) => fmtUGX(v)} />
              <Legend wrapperStyle={{ fontSize: 11.5 }} />
            </PieChart>
          </ResponsiveContainer>
          {totalPayment > 0 && (
            <p style={{ fontSize: 12.5, color: COLORS.inkFaint, marginTop: 6 }}>
              Cash and card payments account for about {fmtPct(cashCardShare)} of the total.
            </p>
          )}
        </Card>
      </div>

      <Card style={{ marginBottom: 24 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
          <SectionTitle sub="Group products with different names (e.g., color/option variants) together as one entry in the best-sellers list">Product Group Management</SectionTitle>
          {!groupFormOpen ? (
            <button onClick={() => setGroupFormOpen(true)} style={btnStyle(COLORS.teal, true)}><Plus size={14} /> Add Group</button>
          ) : (
            <button onClick={() => setGroupFormOpen(false)} style={btnStyle(COLORS.inkFaint, false)}><X size={14} /> Close</button>
          )}
        </div>
        {groupFormOpen && (
          <div style={{ marginTop: 12, padding: 14, borderRadius: 8, background: COLORS.bg, display: "flex", flexDirection: "column", gap: 10 }}>
            <div>
              <label style={{ fontSize: 12.5, color: COLORS.inkFaint, display: "block", marginBottom: 4 }}>Display name (shown for the group)</label>
              <input
                value={groupNameInput} onChange={(e) => setGroupNameInput(e.target.value)} placeholder="e.g., T-shirt"
                style={{ width: "100%", padding: "8px 10px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 13 }}
              />
            </div>
            <div>
              <label style={{ fontSize: 12.5, color: COLORS.inkFaint, display: "block", marginBottom: 4 }}>
                Original product names to merge (comma-separated, exactly as shown in the list)
              </label>
              <input
                value={groupMembersInput} onChange={(e) => setGroupMembersInput(e.target.value)}
                placeholder="e.g., T-shirt, T-shirt with one print extra, T-shirt with more prints"
                style={{ width: "100%", padding: "8px 10px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 13 }}
              />
            </div>
            <button onClick={addProductGroup} style={{ ...btnStyle(COLORS.teal, true), alignSelf: "flex-start" }}>
              <Save size={14} /> Save Group
            </button>
          </div>
        )}
        {productGroups.length > 0 ? (
          <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 14 }}>
            {productGroups.map((g) => {
              const knownNameKeys = new Set([...topProducts.map((p) => p.name), ...allProducts.map((p) => p.name)].map(normalizeProductNameKey));
              const unmatched = g.members.filter((m) => !knownNameKeys.has(normalizeProductNameKey(m)));
              return (
                <div key={g.id} style={{ padding: "6px 0", borderBottom: `1px solid ${COLORS.border}` }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 13 }}>
                    <span><b>{g.canonicalName}</b> ← {g.members.join(", ")}</span>
                    <button onClick={() => removeProductGroup(g.id)} style={{ border: "none", background: "transparent", cursor: "pointer", color: COLORS.clay }}>
                      <Trash2 size={15} />
                    </button>
                  </div>
                  {unmatched.length > 0 && (
                    <p style={{ fontSize: 12, color: COLORS.clay, marginTop: 4 }}>
                      <Info size={12} style={{ verticalAlign: -1, marginRight: 4 }} />
                      Not found in the product list: {unmatched.join(", ")} — please double-check spelling/spacing.
                    </p>
                  )}
                </div>
              );
            })}
          </div>
        ) : (
          !groupFormOpen && <p style={{ fontSize: 12.5, color: COLORS.inkFaint, marginTop: 10 }}>No groups registered.</p>
        )}
      </Card>

      {perf && (
        <Card>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
            <SectionTitle sub="Based on the full POS Data product list · cumulative data (excludes discontinued/sample/service items)">Underperforming Products</SectionTitle>
            {vendorOptions.length > 0 && (
              <select
                value={perfVendorFilter} onChange={(e) => setPerfVendorFilter(e.target.value)}
                style={{ padding: "6px 10px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 12.5, color: COLORS.ink, background: "#fff" }}
              >
                <option value="all">All (every vendor)</option>
                {vendorOptions.map((v) => <option key={v} value={v}>{v} only</option>)}
              </select>
            )}
          </div>
          <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginTop: 8, marginBottom: 14 }}>
            <KpiCard icon={ShoppingBag} label="Sold 1 or Fewer" value={filteredPerf.totalCount > 0 ? fmtPct(filteredPerf.zeroOrOneCount / filteredPerf.totalCount) : "—"} accent={COLORS.clay} note={`${filteredPerf.zeroOrOneCount} / ${filteredPerf.totalCount} item(s)`} />
            <KpiCard icon={ShoppingBag} label="Sold Fewer than 5" value={filteredPerf.totalCount > 0 ? fmtPct(filteredPerf.lessThan5Count / filteredPerf.totalCount) : "—"} accent={COLORS.ochre} note={`${filteredPerf.lessThan5Count} / ${filteredPerf.totalCount} item(s)`} />
          </div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
            {filteredPerf.underperformers.map((name) => (
              <span key={name} style={{ padding: "5px 11px", borderRadius: 20, background: COLORS.bg, color: COLORS.inkSoft, fontSize: 12 }}>{name}</span>
            ))}
            {filteredPerf.underperformers.length === 0 && (
              <span style={{ fontSize: 13, color: COLORS.inkFaint }}>No underperforming products for this vendor.</span>
            )}
          </div>
          <p style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 14 }}>
            <Info size={12} style={{ verticalAlign: -1, marginRight: 4 }} />
            Products with little to no sales are candidates for discontinuation, discounting, or clearing inventory. (Figures are cumulative units sold — splitting by month would be less meaningful, so this is fixed to cumulative.)
            {vendorOptions.length > 0 && " Vendor grouping is based on the POS Data category, so older data without a category may only appear under 'All'."}
          </p>
        </Card>
      )}
    </div>
  );
}

/* ---------------------------------- forecast tab ---------------------------------- */
function ForecastSection({ label, accentColor, forecast, activeMonths }) {
  if (!forecast) {
    return <Card><p style={{ color: COLORS.inkFaint }}>To calculate a forecast, enter at least one month of revenue data in the 'Data Management' tab.</p></Card>;
  }
  const { ytd, avgMonthly, remaining, projectedFlat, trendSeries, annualGoalUGX, requiredAvgRemaining, momGrowth, projectedAchievement } = forecast;
  return (
    <div>
      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 24 }}>
        <KpiCard icon={TrendingUp} label="Projected Annual Revenue" value={fmtCompact(projectedFlat) + " UGX"} accent={accentColor} note={`Assumes the recent ${activeMonths}-month average holds`} />
        <KpiCard icon={Target} label="Projected Annual Goal Achievement" value={fmtPct(projectedAchievement)} accent={projectedAchievement >= 1 ? COLORS.teal : COLORS.clay} note={fmtUGX(annualGoalUGX) + " goal"} />
        <KpiCard icon={Sparkles} label="Monthly Avg. Needed for Goal" value={remaining > 0 ? fmtCompact(requiredAvgRemaining) + " UGX" : "—"} accent={COLORS.ochre} note={remaining > 0 ? `Based on ${remaining} remaining month(s)` : "Full 12 months of data"} />
        <KpiCard icon={ChevronRight} label="Change vs. Previous Month" value={momGrowth === null ? "—" : `${momGrowth >= 0 ? "+" : ""}${(momGrowth * 100).toFixed(1)}%`} accent={momGrowth >= 0 ? COLORS.teal : COLORS.clay} note="Based on the latest month with data" />
      </div>

      <Card style={{ marginBottom: 24 }}>
        <SectionTitle sub={`${label}'s actual revenue (bars) alongside a projection based on the current trend (dashed line)`}>Revenue Forecast</SectionTitle>
        <ResponsiveContainer width="100%" height={320}>
          <ComposedChart data={trendSeries} margin={{ top: 24, right: 10, left: 0, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={COLORS.border} vertical={false} />
            <XAxis dataKey="period" tick={{ fontSize: 12, fill: COLORS.inkSoft }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
            <YAxis tickFormatter={fmtCompact} tick={{ fontSize: 11, fill: COLORS.inkFaint }} axisLine={false} tickLine={false} width={50} />
            <Tooltip contentStyle={tooltipStyle()} formatter={(v, name) => [fmtUGX(v), name === "actual" ? "Actual Revenue" : "Projected Revenue"]} />
            <Legend formatter={(v) => (v === "actual" ? "Actual Revenue" : "Projected Revenue")} wrapperStyle={{ fontSize: 12 }} />
            <Bar dataKey="actual" fill={accentColor} radius={[4, 4, 0, 0]}>
              <LabelList dataKey="actual" content={(props) => <BarValueLabel {...props} />} />
            </Bar>
            <Line dataKey="trend" stroke={COLORS.clay} strokeWidth={2.5} strokeDasharray="6 4" dot={{ r: 3 }} connectNulls>
              <LabelList dataKey="trend" position="top" formatter={(v) => (v > 0 ? fmtCompact(v) : "")} style={{ fontSize: 11, fontWeight: 600, fill: COLORS.clay }} />
            </Line>
          </ComposedChart>
        </ResponsiveContainer>
      </Card>

      <Card>
        <SectionTitle>Forecast Summary</SectionTitle>
        <ul style={{ fontSize: 14, color: COLORS.ink, lineHeight: 1.9, paddingLeft: 20, margin: 0 }}>
          <li>The average monthly revenue over the {activeMonths} month(s) so far is <b>{fmtUGX(avgMonthly)}</b>. (Cumulative: {fmtUGX(ytd)})</li>
          <li>If this trend continues, annual revenue is projected at about <b>{fmtUGX(projectedFlat)}</b>, which is <b>{fmtPct(projectedAchievement)}</b> of the annual goal.</li>
          {remaining > 0 ? (
            <li>To reach the annual goal ({fmtUGX(annualGoalUGX)}), the remaining {remaining} month(s) need average monthly revenue of <b>{fmtUGX(requiredAvgRemaining)}</b>.</li>
          ) : (
            <li>All 12 months of data have been entered.</li>
          )}
        </ul>
      </Card>
    </div>
  );
}

function ForecastTab({ b2cForecast, b2bForecast, data, activeMonths }) {
  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
        <div style={{ width: 10, height: 10, borderRadius: 3, background: COLORS.teal }} />
        <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 22, fontWeight: 700, margin: 0, color: COLORS.ink }}>B2C Forecast (Gift Shop)</h2>
        <span style={{ fontSize: 11, color: COLORS.inkFaint, marginLeft: 2 }}>B2C Forecast</span>
      </div>
      <ForecastSection label="Gift Shop" accentColor={COLORS.teal} forecast={b2cForecast} activeMonths={activeMonths} />

      <div style={{ display: "flex", alignItems: "center", gap: 8, margin: "40px 0 14px" }}>
        <div style={{ width: 10, height: 10, borderRadius: 3, background: COLORS.ochre }} />
        <h2 style={{ fontFamily: "'Fraunces', serif", fontSize: 22, fontWeight: 700, margin: 0, color: COLORS.ink }}>B2B Forecast (Custom Orders)</h2>
        <span style={{ fontSize: 11, color: COLORS.inkFaint, marginLeft: 2 }}>B2B Forecast</span>
      </div>
      <ForecastSection label="Custom Orders" accentColor={COLORS.ochre} forecast={b2bForecast} activeMonths={activeMonths} />

      <p style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 24 }}>
        <Info size={12} style={{ verticalAlign: -1, marginRight: 4 }} />
        Projections are a simple average/linear-trend estimate based on monthly revenue entered so far, and don't account for seasonality or events (markets, sales promotions, etc.).
      </p>
    </div>
  );
}

/* ---------------------------------- data management tab ---------------------------------- */
function DataTab({ data, editMode, draft, startEdit, cancelEdit, saveEdit, resetSample, updateDraftCell, setDraft, onUploadFiles, uploadBusy, uploadLog }) {
  const [confirmReset, setConfirmReset] = useState(false);
  const shown = editMode ? draft : data;
  const allChannelKeys = CHANNEL_KEYS;
  const channelMeta = CHANNEL_META;
  return (
    <div>
      <div style={{
        padding: "10px 14px", borderRadius: 8, background: COLORS.tealSoft, fontSize: 12, color: COLORS.ink, marginBottom: 16,
      }}>
        <Info size={12} style={{ verticalAlign: -1, marginRight: 4 }} />
Anything you upload or edit here is saved <b>only to this browser (this screen)</b>. It's not automatically shared with others. To officially update it for the whole team, either let the daily automation rebuild this page from the Google Sheet, or send the latest Excel file to the person in charge so they can replace this page with a new version.
      </div>
      <Card style={{ marginBottom: 20 }}>
        <SectionTitle sub="When monthly results are ready, re-upload the two original Excel files (Daily Report, Sales Data) as-is. Revenue, visitors, best-sellers, payment mix, day-of-week patterns, and visitor analysis are all automatically recalculated.">
          Update via Excel File
        </SectionTitle>
        <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap", marginTop: 10 }}>
          <label style={{
            display: "flex", alignItems: "center", gap: 8, padding: "10px 18px", borderRadius: 9,
            background: uploadBusy ? COLORS.inkFaint : COLORS.teal, color: "#fff", fontSize: 13.5, fontWeight: 600,
            cursor: uploadBusy ? "default" : "pointer",
          }}>
            <Pencil size={14} />
            {uploadBusy ? "Processing..." : "Choose Excel File(s) (multiple allowed)"}
            <input
              type="file" accept=".xlsx" multiple disabled={uploadBusy} style={{ display: "none" }}
              onChange={(e) => { if (e.target.files && e.target.files.length) onUploadFiles(e.target.files); e.target.value = ""; }}
            />
          </label>
          {data.lastUploadedAt && (
            <span style={{ fontSize: 12, color: COLORS.inkFaint }}>
              Last uploaded: {new Date(data.lastUploadedAt).toLocaleString("en-US")}
            </span>
          )}
        </div>
        {uploadLog && uploadLog.length > 0 && (
          <div style={{ marginTop: 14, padding: 12, borderRadius: 8, background: COLORS.bg, fontSize: 12.5, color: COLORS.ink, lineHeight: 1.8 }}>
            {uploadLog.map((m, i) => <div key={i}>{m}</div>)}
          </div>
        )}
        <ul style={{ fontSize: 12.5, color: COLORS.inkFaint, lineHeight: 1.9, paddingLeft: 20, margin: "14px 0 0" }}>
          <li><b>Sales data file</b> (includes a 'Monthly Sales' sheet, plus 'Details Report'/'POS Data' if present) → updates channel revenue, goal (Goal USD), best-sellers, and payment method mix.</li>
          <li><b>Daily report file</b> (includes a 'Daily Report' sheet) → updates visitors/inquiries/units sold, day-of-week patterns, and visitor origin/type analysis (based on notes).</li>
          <li>Sheet names must <b>not include a year</b> to be auto-recognized (e.g., "Monthly Sales", not "2026 Monthly Sales"). If the year detected in the file is earlier than the current year ({data.currentYear}), it's automatically saved as 'prior-year data'; if it's later, the current year advances to that year and the previous year's data is preserved as prior-year data.</li>
          <li><b>Per-customer custom order details</b> can now be read automatically — either from a separate 'Custom Orders' tab, or from a memo/note on the "Custom Order" row's monthly cell in Monthly Sales, written as "Customer / Category / Product / Qty / Amount" per line (and "LEAD: ..." for a still-open inquiry). If neither is present, add/edit them directly in the 'Custom Orders' tab instead.</li>
        </ul>
      </Card>

      <Card style={{ marginBottom: 20 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12 }}>
          <SectionTitle sub="Use this when you want to edit numbers directly instead of uploading a file.">
            Monthly Performance Data (Direct Edit)
          </SectionTitle>
          <div style={{ display: "flex", gap: 8 }}>
            {!editMode ? (
              <>
                <button onClick={startEdit} style={btnStyle(COLORS.teal, true)}><Pencil size={14} /> Edit</button>
                {!confirmReset ? (
                  <button onClick={() => setConfirmReset(true)} style={btnStyle(COLORS.inkFaint, false)}><RotateCcw size={14} /> Reset to Sample</button>
                ) : (
                  <>
                    <span style={{ fontSize: 12.5, color: COLORS.clay, alignSelf: "center", fontWeight: 600 }}>
                      Everything entered or uploaded so far will be lost. Really reset?
                    </span>
                    <button onClick={() => { resetSample(); setConfirmReset(false); }} style={btnStyle(COLORS.clay, true)}>
                      <RotateCcw size={14} /> Yes, reset
                    </button>
                    <button onClick={() => setConfirmReset(false)} style={btnStyle(COLORS.inkFaint, false)}><X size={14} /> Cancel</button>
                  </>
                )}
              </>
            ) : (
              <>
                <button onClick={saveEdit} style={btnStyle(COLORS.teal, true)}><Save size={14} /> Save</button>
                <button onClick={cancelEdit} style={btnStyle(COLORS.inkFaint, false)}><X size={14} /> Cancel</button>
              </>
            )}
          </div>
        </div>

        <div style={{ display: "flex", gap: 24, margin: "6px 0 18px", flexWrap: "wrap" }}>
          <SettingField label="Gift Shop Monthly Goal (USD)" value={shown.monthlyGoalUSD} editMode={editMode}
            onChange={(v) => setDraft((p) => ({ ...p, monthlyGoalUSD: v === "" ? 0 : Number(v) }))} />
          <SettingField label="Exchange Rate (USD → UGX)" value={shown.exchangeRate} editMode={editMode}
            onChange={(v) => setDraft((p) => ({ ...p, exchangeRate: v === "" ? 0 : Number(v) }))} />
          <SettingField label="ODM (Custom Orders) Annual Goal (UGX)" value={shown.odmAnnualGoalUGX} editMode={editMode}
            onChange={(v) => setDraft((p) => ({ ...p, odmAnnualGoalUGX: v === "" ? 0 : Number(v) }))} />
        </div>

        <div className="sdc-scroll" style={{ overflowX: "auto" }}>
          <table style={{ borderCollapse: "collapse", width: "100%", minWidth: 980, fontSize: 12.5 }}>
            <thead>
              <tr>
                <Th sticky>Month</Th>
                {allChannelKeys.map((k) => <Th key={k}>{channelMeta[k]?.label || k}</Th>)}
                <Th accent={COLORS.ochre}>Visitors</Th>
                <Th accent={COLORS.ochre}>Inquiries</Th>
                <Th accent={COLORS.ochre}>Units Sold</Th>
              </tr>
            </thead>
            <tbody>
              {MONTH_LABELS.map((label, i) => (
                <tr key={label}>
                  <Td sticky bold>{label}</Td>
                  {allChannelKeys.map((k) => (
                    <Td key={k}>
                      <DataCell editMode={editMode} value={(shown.channels[k] || [])[i]}
                        onChange={(v) => updateDraftCell("channel", k, i, v)} />
                    </Td>
                  ))}
                  <Td><DataCell editMode={editMode} value={shown.visits[i]} onChange={(v) => updateDraftCell("visits", null, i, v)} /></Td>
                  <Td><DataCell editMode={editMode} value={shown.contacts[i]} onChange={(v) => updateDraftCell("contacts", null, i, v)} /></Td>
                  <Td><DataCell editMode={editMode} value={shown.sold[i]} onChange={(v) => updateDraftCell("sold", null, i, v)} /></Td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      <Card>
        <SectionTitle>How to Update</SectionTitle>
        <ol style={{ fontSize: 13.5, color: COLORS.ink, lineHeight: 1.9, paddingLeft: 20, margin: 0 }}>
          <li>Easiest method: at the end of each month, upload the two Excel files with the month's results, keeping their original file names, in <b>'Update via Excel File'</b> above. Each file is auto-recognized by its sheet name.</li>
          <li>Per-customer custom order details are picked up automatically from a 'Custom Orders' tab or from a "Customer / Category / Product / Qty / Amount" memo on the Custom Order row — if neither is used, add them directly in the 'Custom Orders' tab instead.</li>
          <li>To quickly fix numbers without an Excel file, use the <b>Edit</b> button below to enter values directly.</li>
          <li>However you update it, the Overview, Visits, Products/Payment, and Forecast tabs are all automatically recalculated as soon as you save.</li>
        </ol>
      </Card>
    </div>
  );
}

function Th({ children, sticky, accent, align }) {
  return (
    <th style={{
      padding: "8px 10px", fontWeight: 600, color: accent || COLORS.inkSoft,
      borderBottom: `1.5px solid ${COLORS.border}`, whiteSpace: "nowrap",
      position: sticky ? "sticky" : "static", left: sticky ? 0 : "auto", background: COLORS.surface,
      textAlign: align || (sticky ? "left" : "right"),
    }}>
      {children}
    </th>
  );
}
function Td({ children, sticky, bold, align }) {
  return (
    <td style={{
      padding: "6px 10px", borderBottom: `1px solid ${COLORS.border}`, textAlign: align || (sticky ? "left" : "right"),
      fontWeight: bold ? 600 : 400, position: sticky ? "sticky" : "static", left: sticky ? 0 : "auto",
      background: COLORS.surface, whiteSpace: "nowrap",
    }}>
      {children}
    </td>
  );
}
function DataCell({ editMode, value, onChange }) {
  if (!editMode) return <span style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{fmtNum(value || 0)}</span>;
  return (
    <input type="number" value={value === 0 ? 0 : value || ""} onChange={(e) => onChange(e.target.value)}
      style={{
        width: 92, padding: "4px 6px", textAlign: "right", border: `1px solid ${COLORS.border}`,
        borderRadius: 6, fontSize: 12.5, background: COLORS.bg, color: COLORS.ink,
      }} />
  );
}
function SettingField({ label, value, editMode, onChange }) {
  return (
    <div>
      <div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginBottom: 4, fontWeight: 600 }}>{label}</div>
      {editMode ? (
        <input type="number" value={value} onChange={(e) => onChange(e.target.value)}
          style={{ width: 120, padding: "6px 8px", border: `1px solid ${COLORS.border}`, borderRadius: 6, fontSize: 13 }} />
      ) : (
        <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 15, fontWeight: 600 }}>{fmtNum(value)}</div>
      )}
    </div>
  );
}
function btnStyle(color, filled) {
  return {
    display: "flex", alignItems: "center", gap: 6, padding: "8px 14px", borderRadius: 8,
    border: filled ? "none" : `1px solid ${COLORS.border}`, cursor: "pointer",
    background: filled ? color : "transparent", color: filled ? "#fff" : COLORS.inkSoft,
    fontSize: 13, fontWeight: 600, fontFamily: "'Inter', sans-serif",
  };
}
