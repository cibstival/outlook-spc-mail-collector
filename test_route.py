# -*- coding: utf-8 -*-
"""route_spc: 파일명의 신보 SPC 이름이 메일 제목 SPC보다 우선(10-23 잘못 분류 회귀 시험). 가상 이름만."""
import ingest_spc as I


class R:
    alias = [("신보2099제11차", "신보2099제11차유동화전문(유)"), ("신보2099", "신보2099제11차유동화전문(유)")]   # 느슨한 별칭이 있어도

    def spc(self, *t):
        t = I.STRIP.sub("", " ".join(x or "" for x in t))
        return next((s for a, s in self.alias if a in t), None)

    new_spc = staticmethod(I.Router.new_spc)


r, subj = R(), "신보2099제11차유동화전문(유)"
ok = [I.route_spc(r, "x", "14_신보2099제3차_자금관리내역(26.06).xlsx", subj, "자금관리내역", peek=False)[:2] == ("신보2099제3차", True),
      I.route_spc(r, "x", "신보2099제11차_자금(26.06).xlsx", subj, "s", peek=False)[0] == subj,
      I.route_spc(r, "x", "자금관리내역(26.06).xlsx", subj, "s", peek=False)[2] == "제목",
      I.route_spc(r, "x", "4. 신보 2099제13차유동화전문(유).pdf", subj, "s", peek=False)[:2] == ("신보2099제13차", True)]
print("결과:", "통과" if all(ok) else f"실패 {ok}")
raise SystemExit(0 if all(ok) else 1)
