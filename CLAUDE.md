# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This is a single-file lunch ordering system (午餐訂購系統) for a small team. The entire app lives in `index.html` — no build step, no dependencies, no package manager.

Open `index.html` directly in a browser to run it.

## Architecture

The app is a vanilla HTML/CSS/JS page that talks to two external services:

- **Google Apps Script** (`SCRIPT_URL` constant): Acts as the backend and database (Google Sheets). All order reads/writes go here via `fetch()`. GET requests use `?action=getData&date=...`; POST requests send `{ action, data, date }` as JSON.
- **ImgBB API** (`IMGBB_API_KEY` constant): Used to host menu images uploaded by the admin.

There is no local state persistence beyond `localStorage` (used only for tracking per-person payment status client-side).

## Key Data Structures

- `orders`: array of `{ id, name, itemName, price }` fetched from the backend each time data changes
- `employeeList`: hardcoded array of employee names (line 432) — update this array to add/remove staff
- `menuImageUrl`: a single string URL pointing to the current day's menu image, stored in Google Sheets

## Admin Password

The admin password is hardcoded as `'10'` (checked via `prompt()`). Required for uploading menus, updating the menu URL, deleting orders, and clearing data.

## Date Model

`currentDate` (YYYY-MM-DD string) drives all data fetches. Viewing past dates disables the order form and clear button. Payment-paid state in localStorage is keyed by `paid_${currentDate}_${name}`.

## Rice Portion & Remarks

These are not separate backend fields. The JS appends them to `itemName` before sending: `排骨飯 (少飯) [加辣]`. The meal chip/autocomplete feature strips these suffixes to extract the base meal name.
