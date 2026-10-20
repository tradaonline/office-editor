import type { Metadata } from "next";
import { PropsWithChildren } from "react";

export const metadata: Metadata = {
  title: "XERP Office",
  robots: { index: false, follow: false },
};

export default function Layout({ children }: PropsWithChildren<unknown>) {
  return <>{children}</>;
}
