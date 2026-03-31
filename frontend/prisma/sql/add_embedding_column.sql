-- Run once in Supabase SQL Editor after prisma migrate (extension vector already enabled).
ALTER TABLE "LectureChunk" ADD COLUMN IF NOT EXISTS embedding vector(3072);

CREATE INDEX IF NOT EXISTS lecturechunk_embedding_hnsw
  ON "LectureChunk"
  USING hnsw (embedding vector_cosine_ops);
