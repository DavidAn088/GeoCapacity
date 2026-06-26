# GeoCapacity GitHub Launch Guide

GitHub profile: https://github.com/DavidAn088

## Option A: Upload manually

1. Go to your GitHub profile.
2. Create a new repository named `GeoCapacity`.
3. Upload every file/folder from this package.
4. Commit the files.

## Option B: Use Git command line

```bash
git clone https://github.com/DavidAn088/GeoCapacity.git
cd GeoCapacity
# copy all GeoCapacity files into this folder
git add .
git commit -m "Launch GeoCapacity app"
git push origin main
```

## Streamlit Community Cloud

1. Open Streamlit Community Cloud.
2. Click **New app**.
3. Repository: `DavidAn088/GeoCapacity`.
4. Main file path: `streamlit_app.py`.
5. Click **Deploy**.
