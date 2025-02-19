import os
from llama_index.llms.azure_openai import AzureOpenAI
from llama_index.core import VectorStoreIndex, SimpleDirectoryReader, get_response_synthesizer, DocumentSummaryIndex
from llama_index.core.indices.document_summary import DocumentSummaryIndexLLMRetriever
from llama_index.core.query_engine import RetrieverQueryEngine
from llama_index.embeddings.azure_openai import AzureOpenAIEmbedding
from llama_index.core.settings import Settings
from secret import *

# Set up LLM and embedding model
llm = AzureOpenAI(
    azure_endpoint=endpoint,
    api_key=api_key,
    api_version=api_version,
    engine=model
)

embed_model = AzureOpenAIEmbedding(
    model=embedding_model,
    deployment_name=embedding_model,
    api_key=api_key,
    azure_endpoint=endpoint,
    api_version=api_version
)

Settings.llm = llm
Settings.embed_model = embed_model



# Define queries for weekly update
queries = {
    "Summary": "Summarize the key activities and achievements from this week's emails. List activities one by one, prioritizing customer interactions over internal activities. Include customer names, company names, and department names if available. When mentioning colleagues, refer to them by their roles or positions rather than names if possible.",
    "Projects": "List the main projects worked on this week, prioritizing customer-related projects. Include customer names, company names, and department names associated with each project if available. Refer to colleagues by their roles or positions rather than names if possible.",
    "Challenges": "List the challenges or obstacles mentioned in the emails, with each challenge on a new line. Prioritize customer-related challenges. Include specific customer names, company names, or department names if they are related to the challenges. Refer to colleagues by their roles or positions rather than names if possible.",
    "Upcoming": "List upcoming tasks or deadlines mentioned for next week, prioritizing customer-related tasks. Include customer names, company names, and department names if they are associated with these tasks. Refer to colleagues by their roles or positions rather than names if possible."
}

# Process documents and generate weekly update
def generate_weekly_update(reader):
    response_synthesizer = get_response_synthesizer(response_mode="tree_summarize", use_async=True)
    doc_summary_index = DocumentSummaryIndex.from_documents(
        reader.load_data(),
        response_synthesizer=response_synthesizer,
        show_progress=True,
    )

    retriever = DocumentSummaryIndexLLMRetriever(
        doc_summary_index,
        choice_batch_size=10,
        choice_top_k=1,
        verbose=True
    )

    query_engine = RetrieverQueryEngine(retriever=retriever, response_synthesizer=response_synthesizer)

    weekly_update = "# Weekly Update\n\n"

    for section, query in queries.items():
        response = query_engine.query(query)
        weekly_update += f"## {section}\n\n{response}\n\n"

    return weekly_update
    
def main():
    # Configure document reader
    reader = SimpleDirectoryReader(input_dir="./docs", recursive=False)
    
    # Generate the weekly update
    weekly_update_content = generate_weekly_update(reader)

    # Print the weekly update to the terminal
    print(weekly_update_content)

    # Save the weekly update to summary.md
    with open('summary.md', 'w', encoding='utf-8') as f:
        f.write(weekly_update_content)

    print("Weekly update has been saved to summary.md")    

if __name__ == "__main__":
    main()
