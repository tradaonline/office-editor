import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Giới thiệu — XERP Office",
  robots: { index: false, follow: false },
};

// Required notices (ONLYOFFICE additional terms 2 and 3): original developer,
// modified version, license, public source. Built from build-time env.
export default function AboutPage() {
  const sourceUrl = process.env.NEXT_PUBLIC_XERP_SOURCE_URL?.trim() || "";
  const changesUrl = sourceUrl ? `${sourceUrl.replace(/\/+$/, "")}/blob/main/XERP_CHANGES.md` : "";

  return (
    <main className="min-h-screen bg-neutral-50 dark:bg-neutral-900 text-neutral-800 dark:text-neutral-200 py-10 px-4 sm:px-6">
      <div className="max-w-2xl mx-auto bg-white dark:bg-neutral-800 border border-neutral-200 dark:border-neutral-700 rounded-xl shadow-xs p-6 sm:p-8 space-y-6">
        <div>
          <h1 className="text-2xl font-bold text-neutral-900 dark:text-neutral-100">
            Giới thiệu XERP Office
          </h1>
          <p className="mt-1 text-sm text-neutral-500 dark:text-neutral-400">
            Bộ công cụ biên tập tài liệu văn phòng trực tiếp trong trình duyệt cho hệ thống XERP.
          </p>
        </div>

        <section className="space-y-3 text-sm leading-relaxed border-t border-neutral-100 dark:border-neutral-700/60 pt-4">
          <h2 className="text-base font-semibold text-neutral-900 dark:text-neutral-100">
            Nhà phát triển gốc &amp; Bản quyền
          </h2>
          <p>
            Bộ ứng dụng biên tập này được xây dựng trên nền tảng ONLYOFFICE, do{" "}
            <strong>Ascensio System SIA</strong> là nhà phát triển gốc.
          </p>
          <p>
            Đây là <strong>bản sửa đổi (modified version)</strong> chạy trong trình duyệt
            qua Next.js cùng ONLYOFFICE web-apps, sdkjs và x2t (WASM), phục vụ tích hợp
            nhúng cho XERP. Bản phát hành này không liên kết hay được chứng thực bởi Ascensio System SIA.
          </p>
          <p>
            Mọi thông báo bản quyền, điều khoản giấy phép và ghi nhận quyền tác giả gốc của
            ONLYOFFICE đi kèm với các thành phần được giữ nguyên không thay đổi.
          </p>
        </section>

        <section className="space-y-3 text-sm leading-relaxed border-t border-neutral-100 dark:border-neutral-700/60 pt-4">
          <h2 className="text-base font-semibold text-neutral-900 dark:text-neutral-100">
            Giấy phép mã nguồn mở
          </h2>
          <p>
            Ứng dụng được phân phối theo giấy phép{" "}
            <a
              href="https://www.gnu.org/licenses/agpl-3.0.html"
              target="_blank"
              rel="noopener noreferrer"
              className="font-medium text-blue-600 dark:text-blue-400 underline hover:text-blue-800 dark:hover:text-blue-300"
            >
              GNU Affero General Public License v3 (AGPL-3.0)
            </a>
            , cùng các điều khoản bổ sung đi kèm từ ONLYOFFICE.
          </p>
        </section>

        <section className="space-y-3 text-sm leading-relaxed border-t border-neutral-100 dark:border-neutral-700/60 pt-4">
          <h2 className="text-base font-semibold text-neutral-900 dark:text-neutral-100">
            Mã nguồn &amp; Lịch sử thay đổi
          </h2>
          <dl className="grid grid-cols-1 sm:grid-cols-3 gap-2 py-2">
            <dt className="font-medium text-neutral-600 dark:text-neutral-400">
              Mã nguồn công khai:
            </dt>
            <dd className="sm:col-span-2">
              {sourceUrl ? (
                <a
                  href={sourceUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="font-medium text-blue-600 dark:text-blue-400 underline break-all hover:text-blue-800 dark:hover:text-blue-300"
                >
                  {sourceUrl}
                </a>
              ) : (
                <span className="text-neutral-500 italic">Đang cập nhật</span>
              )}
            </dd>

            <dt className="font-medium text-neutral-600 dark:text-neutral-400">
              Lịch sử thay đổi:
            </dt>
            <dd className="sm:col-span-2">
              {changesUrl ? (
                <a
                  href={changesUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="font-medium text-blue-600 dark:text-blue-400 underline break-all hover:text-blue-800 dark:hover:text-blue-300"
                >
                  XERP_CHANGES.md
                </a>
              ) : (
                <span className="text-neutral-500 italic">Đang cập nhật</span>
              )}
            </dd>
          </dl>
        </section>
      </div>
    </main>
  );
}
