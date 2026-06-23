# -*- coding: utf-8 -*-
"""Shared Streamlit CSS for ひだまり帳."""
import streamlit as st

def add_css():
    st.markdown("""
    <style>
    .main .block-container {
        padding-top: 1.2rem;
        max-width: 1200px;
    }
    .calendar-title {
        font-size: 1.6rem;
        font-weight: 700;
        margin: 12px 0 10px 0;
        color: #3f3a35;
    }
    .calendar-grid {
        display: grid;
        grid-template-columns: repeat(7, minmax(0, 1fr));
        gap: 0;
        width: 100%;
        border-top: 1px solid #d8d0c4;
        border-left: 1px solid #d8d0c4;
        background: #fffdf8;
    }
    .day-head {
        font-weight: 700;
        text-align: center;
        padding: 8px 4px;
        border-right: 1px solid #d8d0c4;
        border-bottom: 1px solid #d8d0c4;
        background: #f3eee6;
        color: #4b4035;
        font-size: 0.82rem;
        letter-spacing: 0.08em;
    }
    .day-cell {
        min-height: 135px;
        border-right: 1px solid #d8d0c4;
        border-bottom: 1px solid #d8d0c4;
        padding: 8px;
        background: #fffdf8;
        box-sizing: border-box;
        overflow: hidden;
    }
    .blank-cell {
        background: #fffdf8;
    }
    .day-cell-muted {
        min-height: 150px;
        border: 1px solid #eee6dc;
        border-radius: 4px;
        padding: 8px;
        background: #fffdf8;
        color: #aaa;
    }
    .day-num {
        font-size: 2.4rem;
        font-weight: 700;
        line-height: 1;
        margin-bottom: 8px;
        letter-spacing: -1px;
    }
    .write-lines {
        height: 34px;
        margin: 4px 0 6px 0;
        background-image: repeating-linear-gradient(
            to bottom,
            transparent 0px,
            transparent 13px,
            #d9d2c8 14px
        );
        opacity: 0.75;
    }
    .sunday { color: #c0392b; }
    .saturday { color: #1f4e79; }
    .event-line {
        font-size: 0.78rem;
        line-height: 1.35;
        margin: 3px 0;
        padding: 3px 5px;
        border-radius: 5px;
        background: #f6efe6;
        overflow-wrap: anywhere;
        border-left: 3px solid #bfae9b;
    }
    .important {
        background: #ffe9e0;
        border-left: 4px solid #d65a31;
    }
    .small-note {
        color: #7a6a5b;
        font-size: 0.9rem;
    }

    .today-board-card {
        border: 2px solid #d8d0c4;
        border-left: 10px solid #bfae9b;
        background: #fffdf8;
        border-radius: 8px;
        padding: 14px 16px;
        margin: 10px 0;
        box-shadow: none;
    }
    .today-board-card-important {
        border-left: 10px solid #d65a31;
        background: #fff3ee;
    }
    .today-board-main {
        font-size: 1.35rem;
        font-weight: 800;
        line-height: 1.35;
        color: #3f3a35;
    }
    .today-board-time {
        display: inline-block;
        min-width: 76px;
        font-size: 1.45rem;
        font-weight: 900;
        color: #2f2a25;
    }
    .today-board-memo {
        font-size: 1.05rem;
        margin-top: 8px;
        padding-left: 82px;
        color: #4f463d;
    }
    .today-board-sub {
        font-size: 0.9rem;
        margin-top: 8px;
        padding-left: 82px;
        color: #7a6a5b;
    }
    .today-board-badge {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 999px;
        background: #f2eadf;
        font-size: 0.82rem;
        margin-right: 4px;
    }
    .today-summary-box {
        border: 2px solid #d8d0c4;
        background: #f9f5ee;
        border-radius: 8px;
        padding: 12px 14px;
        margin: 8px 0 14px 0;
        font-size: 1.05rem;
        font-weight: 700;
    }


    .hm-calendar-title {
        font-size: 1.45rem;
        font-weight: 800;
        margin: 12px 0 10px;
        color: #2f3437;
    }
    .hm-cal-head {
        text-align: center;
        font-weight: 800;
        padding: 7px 4px;
        border-bottom: 2px solid #d7dde2;
        color: #3f464b;
    }
    .hm-sunday-text { color: #c43d3d; }
    .hm-saturday-text { color: #2466a8; }
    .hm-day-card, .hm-shift-day {
        min-height: 148px;
        border: 1px solid #d9e0e5;
        border-radius: 8px;
        padding: 8px;
        background: #ffffff;
        margin-bottom: 6px;
        overflow: hidden;
    }
    .hm-shift-day { min-height: 132px; }
    .hm-blank { background: #f6f7f8; border-color: #edf0f2; }
    .hm-day-top {
        display: flex;
        align-items: center;
        gap: 6px;
        margin-bottom: 6px;
    }
    .hm-day-number {
        font-size: 1.3rem;
        line-height: 1;
        font-weight: 900;
    }
    .hm-weekday {
        font-size: 0.78rem;
        font-weight: 800;
        color: #69747c;
    }
    .hm-sunday .hm-day-number, .hm-sunday .hm-weekday { color: #c43d3d; }
    .hm-saturday .hm-day-number, .hm-saturday .hm-weekday { color: #2466a8; }
    .hm-today {
        border: 2px solid #2f8f7f;
        background: #f1fbf8;
    }
    .hm-selected {
        box-shadow: inset 0 0 0 2px #3949ab;
    }
    .hm-has-events {
        background: #fffdf6;
    }
    .hm-event-count {
        margin-left: auto;
        padding: 2px 6px;
        border-radius: 999px;
        background: #e9f1ff;
        color: #255a9b;
        font-size: 0.72rem;
        font-weight: 800;
    }
    .hm-event-pill {
        border-left: 4px solid #6d8fbb;
        background: #eef5ff;
        padding: 4px 6px;
        margin: 4px 0;
        border-radius: 6px;
        font-size: 0.78rem;
        line-height: 1.25;
        overflow-wrap: anywhere;
    }
    .hm-more-events, .hm-shift-empty {
        color: #69747c;
        font-size: 0.78rem;
        padding-top: 2px;
    }
    .hm-detail-card {
        border: 1px solid #d9e0e5;
        border-left: 5px solid #4f7fb9;
        border-radius: 8px;
        padding: 10px 12px;
        margin: 10px 0 6px;
        background: #ffffff;
    }
    .hm-detail-title {
        font-weight: 900;
        font-size: 1.05rem;
        color: #2f3437;
    }
    .hm-detail-meta {
        color: #56636b;
        font-size: 0.88rem;
        margin-top: 3px;
    }
    .hm-detail-memo {
        margin-top: 8px;
        color: #3f464b;
        white-space: pre-wrap;
    }
    .hm-shift-line {
        font-size: 0.78rem;
        line-height: 1.35;
        padding: 3px 5px;
        border-radius: 6px;
        background: #f1f6f4;
        margin: 3px 0;
        overflow-wrap: anywhere;
    }

    </style>
    """, unsafe_allow_html=True)


