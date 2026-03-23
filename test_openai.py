from app.modules.chatgpt_client import ChatGPTClient

client = ChatGPTClient()

answer = client.ask(
    system_prompt="あなたは丁寧な日本語アシスタントです。",
    user_prompt="FALCON2向けに自己紹介してください。"
)

print(answer)
