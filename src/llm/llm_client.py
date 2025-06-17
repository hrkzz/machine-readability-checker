import litellm
import streamlit as st  # secrets.toml を使うために必要
from typing import Any

# StreamlitのsecretsからAPIキー取得
litellm.api_key = st.secrets["OPENAI_API_KEY"]

DEFAULT_MODEL = "gpt-4o-mini"
DEFAULT_TEMPERATURE = 0.0
DEFAULT_MAX_TOKENS = 8192

def call_llm(prompt: str, model: str = DEFAULT_MODEL) -> str:
    """
    LiteLLM を通じて OpenAI GPT にプロンプトを送信する。
    モデルは引数で切り替え可能（デフォルト: GPT-4o-mini）。
    """
    try:
        response: Any = litellm.completion(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=DEFAULT_TEMPERATURE,
            max_tokens=DEFAULT_MAX_TOKENS,
        )
        return response["choices"][0]["message"]["content"].strip()
    except Exception as e:
        return f"[ERROR] LLM呼び出しに失敗しました: {e}"
