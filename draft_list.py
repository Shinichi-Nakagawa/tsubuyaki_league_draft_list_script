#!/usr/bin/env python
# -*- coding: utf-8 -*-

__author__ = 'Shinichi Nakagawa'


import os
from datetime import datetime as dt
import xlwt
from sqlalchemy import *
from sqlalchemy.orm import *

from database_config import CONNECTION_TEXT, ENCODING
from tables import *
from stats import Stats


class TsubuyakiLeagueStats(object):

    name = ""
    id = ""
    url = ""

    def __init__(self):
        pass

    @classmethod
    def write_header(cls, row):
        for i, clm in enumerate(cls.HEADER_LIST):
            row.write(i, clm)


class TsubuyakiLeagueStatsPitcher(TsubuyakiLeagueStats):

    HEADER_LIST = (
        "NAME",         # 選手名
        "PlayerID",     # unique key

        # こっから先がゲームのポイント対象
        "IP",           # 投球回
        "W",            # 勝利
        "L",            # 敗戦
        "SV",           # セーブ
        "BB",           # 与四球
        "K",            # 奪三振
        "HLD",          # ホールド
        "ERA",          # 防御率
        "WHIP",         # WHIP
        "QS",           # クオリティースタート
        # 参考の指標
        "GS%",          # 先発回数 / 試合数 * 100
        "K9",           # K/9
        "BB/9",         # BB/9
        "HR/9",         # HR/9
        # Baseball Reference URL
        "url",          # BR
    )
    ip = 0.0
    w = 0
    l = 0
    sv = 0
    bb = 0
    k = 0
    hld = 0
    era = 0.0
    whip = 0.0
    qs = 0
    gs_p = 0
    k9 = 0
    bb9 = 0
    hr9 = 0

    def write_row(self, row):
        row.write(0, self.name)
        row.write(1, self.id)
        row.write(2, self.ip)
        row.write(3, self.w)
        row.write(4, self.l)
        row.write(5, self.sv)
        row.write(6, self.bb)
        row.write(7, self.k)
        row.write(8, self.hld)
        row.write(9, self.era)
        row.write(10, self.whip)
        row.write(11, self.qs)
        row.write(12, self.gs_p)
        row.write(13, self.k9)
        row.write(14, self.bb9)
        row.write(15, self.hr9)
        row.write(16, self.url)


class TsubuyakiLeagueStatsBatter(TsubuyakiLeagueStats):

    HEADER_LIST = (
        "NAME",     # 選手名
        "PlayerID", # unique key

        # こっから先がゲームのポイント対象
        "R",        # 得点
        "H",        # 安打
        "2B",       # 二塁打
        "HR",       # ホームラン
        "RBI",      # 打点
        "SB",       # 盗塁
        "BB",       # 四球
        "A",        # 補殺(assistd)
        "E",        # 失策
        "AVG",      # 打率

        # 選手の打撃能力をサマリーした指標(SABR的なやつ)
        "RC",       # 得点能力(Run Created)
        "OPS",      # OPS(出塁率+長打率, On-base Plus Sluggingの略)
        # Baseball Reference URL
        "url",          # BR
    )
    r = 0
    h = 0
    db = 0
    hr = 0
    rbi = 0
    sb = 0
    bb = 0
    a = 0
    e = 0
    avg = 0.0
    rc = 0.0
    ops = 0.0

    def write_row(self, row):
        row.write(0, self.name)
        row.write(1, self.id)
        row.write(2, self.r)
        row.write(3, self.h)
        row.write(4, self.db)
        row.write(5, self.hr)
        row.write(6, self.rbi)
        row.write(7, self.sb)
        row.write(8, self.bb)
        row.write(9, self.a)
        row.write(10, self.e)
        row.write(11, self.avg)
        row.write(12, self.rc)
        row.write(13, self.ops)
        row.write(14, self.url)


class DraftList(object):

    POSITION_LIST = (
        'P',
        'OF',
        '2B',
        '3B',
        '1B',
        'SS',
        'C',
    )
    BR_URL_FORMAT = "http://www.baseball-reference.com/players/{prefix}/{id}.shtml"

    def __init__(self, session):
        self.session = session

    def find_master_by_year_player(self, playerID):
        query = self.session.query(Master)
        return query.filter_by(playerID=playerID).one()

    def find_batting_by_year_player(self, yearID, playerID):
        query = self.session.query(BattingTotal)
        return query.filter_by(yearID=yearID).filter_by(playerID=playerID)

    def find_fielding_by_year_pos(self, yearID, pos):
        query = self.session.query(Fielding)
        return query.filter_by(yearID=yearID).filter_by(POS=pos).all()

    def find_pitching_by_year_sp(self, yearID, gs=5):
        query = self.session.query(PitchingTotal)
        return query.filter_by(yearID=yearID).filter(PitchingTotal.GS >= gs).all()

    def find_pitching_by_year_p(self, yearID, gs=5):
        query = self.session.query(PitchingTotal)
        return query.filter_by(yearID=yearID).filter(PitchingTotal.GS < gs).all()

    def get_pitching_list(self, pitching_totals):
        """
        投手リストを取得
        :param pitching_list:
        :return:
        """
        pitching_list = []
        for pitching_total in pitching_totals:
            pitching_list.append(self.pitching_stats(pitching_total))
        return pitching_list

    def get_fielding_list(self, fieldings):
        """
        守備位置リストを選手名でchunk
        :param fieldings:
        :return:
        """
        fielding_list = {}
        for fielding in fieldings:
            if fielding.playerID in fielding_list:
                # assistとerrorを足す
                f = fielding_list[fielding.playerID]
                f.a = f.a + fielding.A
                f.e = f.e + fielding.E
            else:
                # 打撃成績を算出してセット
                fielding_list[fielding.playerID] = self.batting_stats(fielding)
        return fielding_list.values()

    def pitching_stats(self, pitching_total):
        """
        投手成績
        :param pitching_total: pitching stats record
        :return: pitching stats
        """
        m = self.find_master_by_year_player(pitching_total.playerID)
        p = TsubuyakiLeagueStatsPitcher()
        p.name = "{first} {last}".format(first=m.nameFirst, last=m.nameLast)
        p.id = pitching_total.playerID
        p.url = self.BR_URL_FORMAT.format(prefix=p.id[0], id=p.id)
        p.ip = Stats.ip(pitching_total.IPouts)
        p.w = pitching_total.W
        p.l = pitching_total.L
        p.sv = 0  # データなし
        p.bb = pitching_total.BB
        p.k = pitching_total.SO
        p.hld = 0  # データなし
        if p.ip > 0.0:
            p.era = Stats.era(pitching_total.ER, p.ip)
            p.whip = Stats.whip(pitching_total.BB, pitching_total.H, p.ip)
            p.gs_p = (pitching_total.GS / pitching_total.G) * 100.0
            p.k9 = Stats.so9(pitching_total.SO, p.ip)
            p.bb9 = Stats.bb9(pitching_total.BB, p.ip)
            p.hr9 = Stats.hr9(pitching_total.HR, p.ip)
        p.qs = 0  # データなし
        return p

    def batting_stats(self, fielding):
        """
        打撃成績
        :param fielding: fielding stats record
        :return: batting stats
        """
        st = TsubuyakiLeagueStatsBatter()
        m = self.find_master_by_year_player(fielding.playerID)
        b = self.find_batting_by_year_player(fielding.yearID, fielding.playerID).one()
        st.name = "{first} {last}".format(first=m.nameFirst, last=m.nameLast)
        st.id = fielding.playerID
        st.url = self.BR_URL_FORMAT.format(prefix=st.id[0], id=st.id)
        st.a = fielding.A
        st.e = fielding.E
        st.r = b.R
        st.h = b.H
        st.db = b._2B
        st.hr = b.HR
        st.rbi = b.RBI
        st.sb = b.SB
        st.bb = b.BB
        if b.AB is not 0:
            single = Stats.single(b.H, b.HR, b._2B, b._3B)
            tb = Stats.tb(
                single,
                b.HR,
                b._2B,
                b._3B
            )
            st.avg = Stats.avg(b.H, b.AB)
            st.rc = Stats.rc(
                b.H,
                b.BB,
                b.HBP,
                b.CS,
                b.GIDP,
                b.SF,
                b.SH,
                b.SB,
                b.SO,
                b.AB,
                b.IBB,
                single,
                b._2B,
                b._3B,
                b.HR
            )
            st.ops = Stats.ops(
                b.H,
                b.BB,
                b.HBP,
                b.AB,
                b.SF,
                tb
            )
        return st

    @classmethod
    def create(cls, values, sheet):

        i = 1
        for value in values:
            row = sheet.row(i)
            value.write_row(row)
            i += 1

# ファイル名フォーマット
FILE_NAME_FORMAT = "{date}_tsubuyaki_league_draft2015.xls"
# 日付フォーマット
DATETIME_FORMAT = "%Y-%m%-d%-H%M%S"
# 対象シーズン
SEASON = 2014


def main(path):

    # Now Date
    now = dt.now()
    # Excel Book
    book = xlwt.Workbook()
    # SQLAlchemy session
    engine = create_engine(CONNECTION_TEXT, encoding=ENCODING)
    Session = sessionmaker(bind=engine, autoflush=True)
    Session.configure(bind=engine)
    session = Session()
    dl = DraftList(session)

    for pos in DraftList.POSITION_LIST:
        if "P" == pos:
            # 投手Stats
            # 中継と先発、別々でリストを作る
            pitching_list = {
                "SP": dl.get_pitching_list(dl.find_pitching_by_year_sp(SEASON)),
                "P": dl.get_pitching_list(dl.find_pitching_by_year_p(SEASON))
            }
            for k, v in pitching_list.items():
                sheet = book.add_sheet(k)
                row = sheet.row(0)
                TsubuyakiLeagueStatsPitcher.write_header(row)
                DraftList.create(v, sheet)
        else:
            # 打者Stats
            sheet = book.add_sheet(pos)
            row = sheet.row(0)
            TsubuyakiLeagueStatsBatter.write_header(row)
            values = dl.get_fielding_list(dl.find_fielding_by_year_pos(SEASON, pos))
            DraftList.create(values, sheet)
    book.save(FILE_NAME_FORMAT.format(date=now.strftime(DATETIME_FORMAT)))


if __name__ == '__main__':
    base_path = os.path.dirname(__file__)
    main(base_path)
