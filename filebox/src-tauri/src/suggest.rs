use crate::model::{ext_of, stem_of, tokenize, Favorite, MoveRecord, Rule, Suggestion};
use crate::rules::rule_favorites;
use std::collections::HashMap;

/// 확장자만 일치하는 약한 신호로 추천을 띄우기 위해 필요한 최소 이동 횟수.
/// xlsx/pdf처럼 흔한 확장자에서 한 번의 이동이 모든 동종 파일로 번지는 것을 막는다.
const EXT_ONLY_MIN_MOVES: i64 = 3;

/// 파일명 하나에 대해 즐겨찾기 추천 목록(점수 내림차순, 최대 `limit`개)을 계산한다.
///
/// 점수 구성:
/// - 규칙이 직접 지정한 즐겨찾기: +40 (첫 규칙), 이후 규칙은 +30
/// - 파일명 토큰 일치: 기록마다 토큰당 +4 (기록당 최대 +12)
/// - 확장자 일치: 같은 확장자를 그 즐겨찾기로 `EXT_ONLY_MIN_MOVES`회 이상 보냈을 때만 +6
/// - 사용 빈도: 이미 다른 신호로 점수를 받은 즐겨찾기에 한해 기록당 +1 (최대 +5)
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

    // 즐겨찾기별로 학습 신호를 따로 모은다.
    let mut token_score: HashMap<String, i64> = HashMap::new();
    let mut ext_moves: HashMap<String, i64> = HashMap::new();
    let mut usage: HashMap<String, i64> = HashMap::new();

    for rec in records {
        *usage.entry(rec.favorite_id.clone()).or_default() += 1;

        if !ext.is_empty() && rec.ext == ext {
            *ext_moves.entry(rec.favorite_id.clone()).or_default() += 1;
        }
        let overlap = rec.tokens.iter().filter(|t| tokens.contains(t)).count() as i64;
        if overlap > 0 {
            *token_score.entry(rec.favorite_id.clone()).or_default() += (overlap * 4).min(12);
        }
    }

    // 파일명이 겹치는 신호는 그 자체로 유효하다.
    for (fav_id, s) in token_score {
        *scores.entry(fav_id).or_default() += s;
    }
    // 확장자만 겹치는 신호는 반복 확인된 경우에만 인정한다.
    for (fav_id, moves) in ext_moves {
        if moves >= EXT_ONLY_MIN_MOVES {
            *scores.entry(fav_id).or_default() += 6;
        }
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
            id: format!("{ext}-{fav}-{}", tokens.join("_")),
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

    /// 확장자만 같고 파일명이 전혀 겹치지 않으면, 반복 이동 전까지는 추천하지 않는다.
    #[test]
    fn ext_only_match_needs_repetition() {
        let favorites = vec![fav("f_proj", "국책법카")];
        let few = vec![
            rec("xlsx", &["국책과제", "실적이체"], "f_proj"),
            rec("xlsx", &["국책과제", "실적이체", "2026"], "f_proj"),
        ];
        let out = suggest(&[], &few, &favorites, "월간_매출집계.xlsx", 3);
        assert!(out.is_empty(), "2회 이동만으로 무관한 xlsx에 추천이 떠서는 안 된다");

        let mut many = few.clone();
        many.push(rec("xlsx", &["국책과제", "실적이체", "3분기"], "f_proj"));
        let out = suggest(&[], &many, &favorites, "월간_매출집계.xlsx", 3);
        assert_eq!(out.len(), 1, "3회 이상이면 확장자 신호를 인정한다");
    }

    /// 파일명이 겹치면 이동 횟수와 무관하게 첫 번째부터 추천한다.
    #[test]
    fn token_match_works_from_first_move() {
        let favorites = vec![fav("f_proj", "국책법카")];
        let records = vec![rec("xlsx", &["국책과제", "실적이체"], "f_proj")];
        let out = suggest(&[], &records, &favorites, "국책과제_실적이체_8월.xlsx", 3);
        assert_eq!(out.len(), 1);
        assert_eq!(out[0].favorite.id, "f_proj");
    }
}
