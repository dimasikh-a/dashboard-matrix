import { NextResponse } from "next/server";
import { Prisma } from "@prisma/client";
import { prisma } from "@/lib/prisma";

type ActivityPayload = { id: string; data: Record<string, string> };

function fromDatabase(activity: {
  id: string; nama: string; jenisKegiatan: string; kasubid: string; kabid: string; kasubidKabidBidangLain: string;
  sekban: string; kaban: string; keteranganTambahan: string; status: string; tanggalSelesai: string;
  dasarRegulasi: string; catatan: string; link: string;
}) {
  return {
    id: activity.id,
    data: {
      Nama: activity.nama,
      "Jenis Kegiatan": activity.jenisKegiatan,
      Kasubid: activity.kasubid,
      Kabid: activity.kabid,
      "Kasubid/Kabid Bidang lain": activity.kasubidKabidBidangLain,
      Sekban: activity.sekban,
      Kaban: activity.kaban,
      "Keterangan Tambahan": activity.keteranganTambahan,
      Status: activity.status,
      "Tanggal Selesai": activity.tanggalSelesai,
      "Perda/Perbup/Kepbup/Perkaban/KepKaban dll": activity.dasarRegulasi,
      Catatan: activity.catatan,
      LINK: activity.link,
    },
  };
}

function toDatabase(data: Record<string, string>): Omit<Prisma.ActivityUncheckedCreateInput, "id"> {
  return {
    nama: data.Nama || "",
    jenisKegiatan: data["Jenis Kegiatan"] || "",
    kasubid: data.Kasubid || "",
    kabid: data.Kabid || "",
    kasubidKabidBidangLain: data["Kasubid/Kabid Bidang lain"] || "",
    sekban: data.Sekban || "",
    kaban: data.Kaban || "",
    keteranganTambahan: data["Keterangan Tambahan"] || "",
    status: data.Status || "Direncanakan",
    tanggalSelesai: data["Tanggal Selesai"] || "",
    dasarRegulasi: data["Perda/Perbup/Kepbup/Perkaban/KepKaban dll"] || "",
    catatan: data.Catatan || "",
    link: data.LINK || "",
  };
}

function recordFor(activity: ActivityPayload) {
  return { ...toDatabase(activity.data), id: activity.id };
}

export async function GET() {
  const activities = await prisma.activity.findMany({ orderBy: { nama: "asc" } });
  return NextResponse.json(activities.map(fromDatabase));
}

export async function POST(request: Request) {
  const activity = (await request.json()) as ActivityPayload;
  if (!activity.id || !activity.data) return NextResponse.json({ message: "Data tidak valid." }, { status: 400 });
  const saved = await prisma.activity.upsert({
    where: { id: activity.id },
    update: toDatabase(activity.data),
    create: recordFor(activity),
  });
  return NextResponse.json(fromDatabase(saved));
}

export async function PATCH(request: Request) {
  const activity = (await request.json()) as ActivityPayload;
  if (!activity.id || !activity.data) return NextResponse.json({ message: "Data tidak valid." }, { status: 400 });
  const saved = await prisma.activity.update({ where: { id: activity.id }, data: toDatabase(activity.data) });
  return NextResponse.json(fromDatabase(saved));
}

export async function DELETE(request: Request) {
  const { id } = (await request.json()) as { id?: string };
  if (!id) return NextResponse.json({ message: "ID tidak tersedia." }, { status: 400 });
  await prisma.activity.delete({ where: { id } });
  return new NextResponse(null, { status: 204 });
}

export async function PUT(request: Request) {
  const { activities } = (await request.json()) as { activities?: ActivityPayload[] };
  if (!activities?.every((activity) => activity.id && activity.data)) return NextResponse.json({ message: "Data tidak valid." }, { status: 400 });
  await prisma.$transaction([
    prisma.activity.deleteMany(),
    prisma.activity.createMany({ data: activities.map(recordFor) }),
  ]);
  return NextResponse.json({ count: activities.length });
}
