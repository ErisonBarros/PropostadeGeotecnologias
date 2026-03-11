import asyncio
from notebooklm import NotebookLMClient

async def main():
    async with await NotebookLMClient.from_storage() as client:
        # Create notebook
        nb = await client.notebooks.create("Quantum Computing Podcast")
        
        # Add source
        print("Adding source...")
        await client.sources.add_url(nb.id, "https://en.wikipedia.org/wiki/Quantum_computing", wait=True)

        # Generate audio
        print("Generating podcast...")
        status = await client.artifacts.generate_audio(nb.id, instructions="Create a podcast about quantum computing")
        await client.artifacts.wait_for_completion(nb.id, status.task_id, timeout=1200)
        
        print("Downloading podcast...")
        try:
            await client.artifacts.download_audio(nb.id, "quantum_podcast.mp3")
        except FileExistsError:
            import os
            os.remove("quantum_podcast.mp3")
            await client.artifacts.download_audio(nb.id, "quantum_podcast.mp3")

        # Generate quiz
        print("Generating quiz...")
        status = await client.artifacts.generate_quiz(nb.id)
        await client.artifacts.wait_for_completion(nb.id, status.task_id, timeout=1200)
        
        print("Downloading quiz...")
        try:
            await client.artifacts.download_quiz(nb.id, "quantum_quiz.md", output_format="markdown")
        except FileExistsError:
            import os
            os.remove("quantum_quiz.md")
            await client.artifacts.download_quiz(nb.id, "quantum_quiz.md", output_format="markdown")

        print("Done!")

asyncio.run(main())
