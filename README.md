# Yuri: Framework for Robust Legal Contract Drafting with Domain-Driven Design

## Data preparation
1. Make sure you have git-lfs installed (https://git-lfs.com) by running `git lfs install`.
2. When prompted for a password, use an access token with write permissions. Generate one from your settings: https://huggingface.co/settings/tokens.
3. Download the datasets:
```shell
git clone https://huggingface.co/datasets/Aborevsky01/Yuri_Legal_Contracts
```
## Preparing the environment
Download all necessary libraries and dependencies:
```shell
pip install -r requirements.txt
```
## Web service deployment
```shell
python -m streamlit run streamlit_app.py
```
