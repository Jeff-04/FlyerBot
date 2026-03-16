import pandas as pd
import telebot
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton
from io import BytesIO
from datetime import datetime
from textwrap import dedent

API_TOKEN = '8442164414:AAE1UbEpL0NeWDfwJg9iuPqEnLgChgPdz88'
CHAT_ID = 1464789247
bot = telebot.TeleBot(API_TOKEN)
user_state = {}


def transform_jadwal(df):
    """Mengubah dataframe menjadi format kolom yang diinginkan"""
    
    hasil = pd.DataFrame()

    # Nomor urut
    hasil["NO"] = range(1, len(df) + 1)

    # Nama pengajar
    hasil["NAMA"] = df["nama_pengajar"]

    # Jam mulai dan selesai
    hasil["JAM"] = df["jam_awal"]
    hasil["JAM_2"] = df["jam_akhir"]

    # Kelas
    hasil["KELAS"] = df["nama_kelas"]

    # Mapel
    hasil["MAPEL"] = df["mapel_yang_diajarkan"]

    # Jenis (contoh aturan sederhana)
    hasil["JENIS"] = df["status_pengajar"].apply(
        lambda x: "TST" if x == "PF" else "KBM"
    )

    hasil["JENIS"] = "KBM"

    # Ruang kelas (contoh ambil bagian akhir dari nama_kelas)
    hasil["R KELAS"] = "-"
    hasil["Unit"] = df["nama_gedung"]

    return hasil


def kirim_dataframe(df):

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

    buffer.seek(0)

    bot.send_document(
        CHAT_ID,
        buffer,
        visible_file_name=f"jadwal_{format_tanggal_indonesia(datetime.now())}.xlsx"
    )

def format_tanggal_indonesia(date: datetime) -> str:
    hari = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
    bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
             "Juli", "Agustus", "September", "Oktober", "November", "Desember"]

    nama_hari = hari[date.weekday()]           # 0=Senin ... 6=Minggu
    nama_bulan = bulan[date.month - 1]         # month 1-12
    return f"{nama_hari}, {date.day} {nama_bulan} {date.year}"


def generate_reminder(df):
    # -----------------------------------------
    # Bagian 1 — Pembuka
    # -----------------------------------------
    nama_pengajar = df['nama_pengajar'].unique()[0]
    unit_ngajar = ", ".join(dict.fromkeys(df['nama_gedung'].tolist()))
    tanggal_hari_ini = format_tanggal_indonesia(datetime.now())

    caption = dedent(f"""
        Selamat Siang Bapak/Ibu {str(nama_pengajar).capitalize()}.

        Hari ini {tanggal_hari_ini} ada jadwal KBM di {unit_ngajar}.
    """).strip() + "\n\n"

    # -----------------------------------------
    # Bagian 2 — Daftar Jadwal
    # -----------------------------------------
    jadwal_items = []

    for _, row in df.iterrows():
        item = dedent(f"""
            📍 *{row['nama_kelas']}*
            ⏰ {row["jam_awal"]} – {row["jam_akhir"]}
            📘 {row['mapel_yang_diajarkan']}
        """).strip()

        jadwal_items.append(item)

    caption += "\n\n".join(jadwal_items)
    caption += "\n\n---\n\n"

    # -----------------------------------------
    # Bagian 3 — Penutup & Reminder
    # -----------------------------------------
    caption_penutup = dedent("""
        ⚠️ Bapak Ibu, KBM dan TST hari ini WAJIB DI REALISASIKAN LANGSUNG SETELAH KEGIATAN SELESAI. Lebih dari pukul 20.00 belum direalisasi maka honor tidak terhitung oleh sistem ⚠️

        📌 *Note:* 

        - Di mohon hadir maksimal 30 menit sebelum jam KBM dimulai
        - Untuk mengantisipasi adanya perubahan jadwal, silakan cek jadwal di *GO TIM* hari ini
        - Terima kasih untuk Bapak/Ibu pengajar yang sudah melakukan penyesuaian BAH dan tidak cancel jadwal KBM.

        Terima kasih 🙏
    """).strip()

    caption += caption_penutup

    return caption


@bot.message_handler(commands=['start'])
def start(message):
    bot.reply_to(message, "Halo! Selamat datang di bot Flyer 🎉")
    menu(message.chat.id)

def menu(chat_id):
    markup = InlineKeyboardMarkup()
    markup.add(InlineKeyboardButton("Flyer TST", callback_data="tst"))
    markup.add(InlineKeyboardButton("Reminder", callback_data="reminder"))
    markup.add(InlineKeyboardButton("TST CSV", callback_data="tst-csv"))
    markup.add(InlineKeyboardButton("Kirim Dokumentasi", callback_data="dokumentasi"))
    markup.add(InlineKeyboardButton("❌ Tutup", callback_data="tutup"))
    bot.send_message(chat_id, "Silakan pilih menu:", reply_markup=markup)


# Handle CSV
@bot.message_handler(content_types=['document'])
def handle_docs(message):
    try:
        # Pastikan file berformat .xlsx
        file_name = message.document.file_name
        if file_name.endswith('.xlsx'):
            # Ambil file dari Telegram
            file_id = message.document.file_id
            file_info = bot.get_file(file_id)
            downloaded_file = bot.download_file(file_info.file_path)

            # Baca file Excel langsung dari memory (BytesIO)
            excel_data = BytesIO(downloaded_file)
            df = pd.read_excel(excel_data)  # <-- Perhatikan ini

            df.to_csv(f"CSV REMINDER/KBM - {format_tanggal_indonesia(datetime.now())}.csv")

            bot.reply_to(message, f"✅ File Excel diterima!")
            print(df.columns)

            df['nama_pengajar'] = df['nama_pengajar'].astype(str)
            df['nama_pengajar'] = df['nama_pengajar'].str.strip()
            df['nama_pengajar'] = df['nama_pengajar'].str.replace(r'\s+', ' ', regex=True)

            for x in df["nama_pengajar"].unique():
                bot.send_message(message.chat.id, generate_reminder(df[df['nama_pengajar'] == x]))
            
            print("Sudah Selesai Semua Terima Kasih !")
            bot.send_message(message.chat.id, "Sudah Selesai Semua Terima Kasih ☺️☺️☺️")

            df_baru = transform_jadwal(df)
            kirim_dataframe(df_baru)

        else:
            bot.reply_to(message, f"✅ File CSV diterima!")



    except Exception as e:
        bot.reply_to(message, f"❌ Gagal membaca file:\n{e}")\


print("Bot jalan...")
bot.infinity_polling()
