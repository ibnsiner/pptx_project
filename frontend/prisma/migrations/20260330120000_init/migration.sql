-- CreateTable
CREATE TABLE "Lecture" (
    "id" TEXT NOT NULL,
    "title" TEXT NOT NULL,
    "description" TEXT,
    "thumbnail" TEXT,
    "author" TEXT NOT NULL,
    "tags" TEXT[] DEFAULT ARRAY[]::TEXT[],
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "Lecture_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "Slide" (
    "id" TEXT NOT NULL,
    "slideNumber" INTEGER NOT NULL,
    "lectureId" TEXT NOT NULL,
    "elements" JSONB NOT NULL,

    CONSTRAINT "Slide_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "LectureChunk" (
    "id" TEXT NOT NULL,
    "lectureId" TEXT NOT NULL,
    "slideId" TEXT,
    "slideNumber" INTEGER NOT NULL,
    "chunkIndex" INTEGER NOT NULL,
    "content" TEXT NOT NULL,
    "tokenCount" INTEGER,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "LectureChunk_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE INDEX "LectureChunk_lectureId_idx" ON "LectureChunk"("lectureId");

-- CreateIndex
CREATE UNIQUE INDEX "LectureChunk_lectureId_slideNumber_key" ON "LectureChunk"("lectureId", "slideNumber");

-- AddForeignKey
ALTER TABLE "Slide" ADD CONSTRAINT "Slide_lectureId_fkey" FOREIGN KEY ("lectureId") REFERENCES "Lecture"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "LectureChunk" ADD CONSTRAINT "LectureChunk_lectureId_fkey" FOREIGN KEY ("lectureId") REFERENCES "Lecture"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "LectureChunk" ADD CONSTRAINT "LectureChunk_slideId_fkey" FOREIGN KEY ("slideId") REFERENCES "Slide"("id") ON DELETE SET NULL ON UPDATE CASCADE;
