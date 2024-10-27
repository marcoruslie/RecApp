<template>
	<div class="p-8">
		<form @submit.prevent="generateDocx">
			<div class="mb-4">
				<label class="block text-sm font-medium">Name</label>
				<input
					type="text"
					v-model="name"
					class="input"
					required />
			</div>

			<div class="mb-4">
				<label class="block text-sm font-medium">Email</label>
				<input
					type="email"
					v-model="email"
					class="input"
					required />
			</div>

			<div class="mb-4">
				<label class="block text-sm font-medium">Message</label>
				<textarea
					v-model="message"
					class="input"
					rows="4"
					required></textarea>
			</div>

			<button
				type="submit"
				class="btn">
				Generate DOCX
			</button>
		</form>
	</div>
</template>

<script setup>
	import { Packer, Document, Paragraph, TextRun } from "docx"

	const name = ref("")
	const email = ref("")
	const message = ref("")

	const generateDocx = async () => {
		// 1. Create a new document
		const doc = new Document({
			sections: [
				{
					properties: {},
					children: [
						new Paragraph({
							children: [new TextRun(`Name: ${name.value}`)],
						}),
						new Paragraph({
							children: [new TextRun(`Email: ${email.value}`)],
						}),
						new Paragraph({
							children: [new TextRun(`Message:`)],
						}),
						new Paragraph({
							children: [new TextRun(`${message.value}`)],
						}),
					],
				},
			],
		})

		// 2. Convert the document to a Blob
		const blob = await Packer.toBlob(doc)

		// 3. Create a download link for the generated .docx file
		const url = URL.createObjectURL(blob)
		const link = document.createElement("a")
		link.href = url
		link.download = "document.docx"
		link.click()

		// 4. Clean up the URL object
		URL.revokeObjectURL(url)
	}
</script>

<style>
	.input {
		width: 100%;
		padding: 8px;
		border: 1px solid #ccc;
		border-radius: 4px;
	}

	.btn {
		background-color: #006780;
		color: white;
		padding: 10px 20px;
		border: none;
		border-radius: 4px;
		cursor: pointer;
	}
</style>
