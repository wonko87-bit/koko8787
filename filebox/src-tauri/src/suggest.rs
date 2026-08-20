use crate::model::{ext_of, stem_of, tokenize, Favorite, MoveRecord, Rule, Suggestion};
use crate::rules::rule_favorites;
use std::collections::HashMap;

/// 파일명 하나에 대해 즐겨찾기 추천 목록(점수 내림차순, 최대 `limit`개)을 계산한다.
///
/// 점수 구성:
/// - 규칙이 직접 지정한 즐겨찾기: +40 (첫 규칙), 이후 규칙은 +30
/// - 학습 기록: 확장자 일치 +6, 파일명 토큰 일치 토큰당 +4 (기록당 최대 +12)
/// - 학습 기록의 단순 사용 빈도: 기록당 +1 (최대 +5)
pub fn suggest(
    rules: &[Rule],
    records: &[MoveRecord],
    favorites: &[Favorite],
    file_name: &str,
    limit: usize,
) -> Vec<Suggestion> {
    let mut scores: HashMap<String, i64> = HashMap::new();

    for (i, fav_id) in rule_favorites(rules, file_name).iter().enumerate() {
        *scores.entry(fav_id.clone()).or_default() += if i == 0 { 40 } else { 30 };
    }

    let ext = ext_of(file_name);
    let tokens = tokenize(&stem_of(file_name));
    let mut usage: HashMap<String, i64> = HashMap::new();

    for rec in records {
        let mut s = 0i64;
        if !ext.is_empty() && rec.ext == ext {
            s += 6;
        }
        let overlap = rec.tokens.iter().filter(|t| tokens.contains(t)).count() as i64;
        s += (overlap * 4).min(12);
        if s > 0 {
            *scores.entry(rec.favorite_id.clone()).or_default() += s;
        }
        *usage.entry(rec.favorite_id.clone()).or_default() += 1;
    }
    for (fav_id, count) in usage {
        if let Some(s) = scores.get_mut(&fav_id) {
            *s += count.min(5);
        }
    }

    let mut result: Vec<Suggestion> = favorites
        .iter()
        .filter_map(|f| {
            scores
                .get(&f.id)
                .filter(|s| **s > 0)
                .map(|s| Suggestion { favorite: f.clone(), score: *s })
        })
        .collect();
    result.sort_by(|a, b| b.score.cmp(&a.score));
    result.truncate(limit);
    result
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::Rule;

    fn fav(id: &str, name: &str) -> Favorite {
        Favorite { id: id.into(), name: name.into(), path: format!("/dest/{id}").into() }
    }

    fn rec(ext: &str, tokens: &[&str], fav: &str) -> MoveRecord {
        MoveRecord {
            ext: ext.into(),
            tokens: tokens.iter().map(|s| s.to_string()).collect(),
            favorite_id: fav.into(),
            at: 0,
        }
    }

    #[test]
    fn rule_favorite_beats_learned() {
        let rules = vec![Rule {
            id: "r".into(),
            name: "인보이스".into(),
            extensions: vec!["pdf".into()],
            keywords: vec![],
            category: None,
            favorite_id: Some("f_docs".into()),
        }];
        let records = vec![rec("pdf", &["invoice"], "f_other")];
        let favorites = vec![fav("f_docs", "문서함"), fav("f_other", "기타함")];
        let out = suggest(&rules, &records, &favorites, "invoice_2026.pdf", 3);
        assert_eq!(out[0].favorite.id, "f_docs");
        assert!(out.len() >= 2);
    }

    #[test]
    fn learned_tokens_rank_destination() {
        let records = vec![
            rec("pdf", &["회의록"], "f_meeting"),
            rec("pdf", &["회의록", "주간"], "f_meeting"),
            rec("pdf", &["사진"], "f_photo"),
        ];
        let favorites = vec![fav("f_meeting", "회의자료"), fav("f_photo", "사진")];
        let out = suggest(&[], &records, &favorites, "주간_회의록_0819.pdf", 3);
        assert_eq!(out[0].favorite.id, "f_meeting");
    }

    #[test]
    fn no_signal_returns_empty() {
        let favorites = vec![fav("f1", "아무곳")];
        let out = suggest(&[], &[], &favorites, "unknown.bin", 3);
        assert!(out.is_empty());
    }

    #[test]
    fn deleted_favorite_not_suggested() {
        let records = vec![rec("pdf", &["report"], "f_gone")];
        let out = suggest(&[], &records, &[], "report.pdf", 3);
        assert!(out.is_empty());
    }
}
