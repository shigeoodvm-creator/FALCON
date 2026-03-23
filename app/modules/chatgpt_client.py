"""
FALCON2 ChatGPT クライアント
- OpenAI API を安全にラップする
- UIや業務ロジックから直接 API を触らせない
"""

import json
from pathlib import Path
from openai import OpenAI


class ChatGPTClient:
    def __init__(self):
        config_path = Path("config/openai.json")

        if not config_path.exists():
            raise FileNotFoundError("config/openai.json が見つかりません")

        config = json.loads(config_path.read_text(encoding="utf-8"))

        self.model = config.get("model", "gpt-4.1-mini")
        self.temperature = config.get("temperature", 0.3)

        self.client = OpenAI(
            api_key=config["api_key"],
            timeout=15.0,
        )

    def ask(self, system_prompt: str, user_prompt: str) -> str:
        """
        ChatGPT に問い合わせて、テキストだけを返す
        """
        response = self.client.chat.completions.create(
            model=self.model,
            temperature=self.temperature,
            messages=[
                {
                    "role": "system",
                    "content": system_prompt,
                },
                {
                    "role": "user",
                    "content": user_prompt,
                },
            ],
        )

        return response.choices[0].message.content
