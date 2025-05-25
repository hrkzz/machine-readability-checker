import litellm
import streamlit as st  # secrets.toml を使うために必要
from typing import Any

# StreamlitのsecretsからAPIキー取得
litellm.api_key = st.secrets["OPENAI_API_KEY"]

DEFAULT_MODEL = "gpt-4o-mini"
DEFAULT_TEMPERATURE = 0.0
DEFAULT_MAX_TOKENS = 8192

def call_llm(prompt: str) -> str:
    """
    LiteLLM を通じて OpenAI GPT-4 にプロンプトを送信する。
    モデル設定はこのモジュール内で管理する。
    """
    try:
        # CustomStreamWrapper の型を明示的に Any にして、インデックスアクセスを許可
        response: Any = litellm.completion(
            model=DEFAULT_MODEL,
            messages=[{"role": "user", "content": prompt}],
            temperature=DEFAULT_TEMPERATURE,
            max_tokens=DEFAULT_MAX_TOKENS,
        )
        return response["choices"][0]["message"]["content"].strip()
    except Exception as e:
        return f"[ERROR] LLM呼び出しに失敗しました: {e}"
