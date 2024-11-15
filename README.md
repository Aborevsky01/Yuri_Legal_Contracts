# Yuri: Framework for Robust Legal Contract Drafting with Domain-Driven Design
## Preparing the environment
Download all necessary libraries and dependencies:
```shell
pip install -r requirements.txt
```
## Data preparation
1. Make sure you have git-lfs installed (https://git-lfs.com) by running `git lfs install`.
2. When prompted for a password, use an access token with write permissions. Generate one from your settings: https://huggingface.co/settings/tokens.
3. Download the datasets:
```shell
git clone https://huggingface.co/datasets/Aborevsky01/Yuri_Legal_Contracts
```
## Web service deployment
```shell
python3 -m streamlit run streamlit_app.py
```
it is possible to specify a custom port http://localhost:8502/:
```shell
python3 -m streamlit run streamlit_app.py --server.port 8502
```
