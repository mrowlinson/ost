//! Microsoft Graph API: joined teams and channels

use anyhow::{bail, Context, Result};
use serde::Deserialize;

use super::client::TeamsClient;

#[derive(Debug, Deserialize)]
struct TeamsResponse {
    value: Vec<Team>,
}

#[derive(Debug, Deserialize)]
struct Team {
    id: String,
    #[serde(rename = "displayName")]
    display_name: Option<String>,
}

#[derive(Debug, Deserialize)]
struct ChannelsResponse {
    value: Vec<Channel>,
}

#[derive(Debug, Deserialize)]
struct Channel {
    id: String,
    #[serde(rename = "displayName")]
    display_name: Option<String>,
}

/// List joined teams and channels (prints to stdout).
pub async fn list_teams() -> Result<()> {
    let client = TeamsClient::new().await?;
    let teams = list_teams_data(&client).await?;

    println!("\nTeams and Channels:");
    println!("{:-<60}", "");

    if teams.is_empty() {
        println!("  (no teams found)");
        return Ok(());
    }

    for team in &teams {
        println!("Team: {} ({} channels)", team.name, team.channels.len());
        for ch in &team.channels {
            println!("  {:<30} {}", ch.name, ch.id);
        }
        println!();
    }

    Ok(())
}

// ---------------------------------------------------------------------------
// Data-returning API functions for TUI integration
// ---------------------------------------------------------------------------

/// Team metadata for TUI display.
pub struct TeamInfo {
    pub id: String,
    pub name: String,
    pub channels: Vec<ChannelInfo>,
}

/// Channel metadata for TUI display.
pub struct ChannelInfo {
    pub id: String,
    pub name: String,
}

/// List joined teams with their channels and return structured data.
pub async fn list_teams_data(client: &TeamsClient) -> Result<Vec<TeamInfo>> {
    tracing::debug!("Fetching joined teams...");
    let resp = client.graph_get("/me/joinedTeams").await?;
    let teams: TeamsResponse = resp
        .json()
        .await
        .context("Failed to parse joinedTeams response")?;

    let mut result = Vec::new();

    for team in &teams.value {
        let team_name = team.display_name.as_deref().unwrap_or(&team.id).to_string();
        tracing::debug!("Fetching channels for team: {} ({})", team_name, team.id);

        let path = format!("/teams/{}/channels", team.id);
        let channels = match client.graph_get(&path).await {
            Ok(resp) => {
                let channels_resp: ChannelsResponse = resp
                    .json()
                    .await
                    .context("Failed to parse channels response")?;
                channels_resp
                    .value
                    .into_iter()
                    .map(|ch| ChannelInfo {
                        name: ch.display_name.unwrap_or_else(|| ch.id.clone()),
                        id: ch.id,
                    })
                    .collect()
            }
            Err(e) => {
                tracing::warn!("Failed to fetch channels for {}: {:#}", team_name, e);
                Vec::new()
            }
        };

        result.push(TeamInfo {
            id: team.id.clone(),
            name: team_name,
            channels,
        });
    }

    Ok(result)
}

/// Reject ids that would break out of the Graph path segment.
fn check_id(what: &str, id: &str) -> Result<()> {
    if id.trim().is_empty() {
        bail!("empty {}", what);
    }
    if id.contains('/')
        || id.contains('?')
        || id.contains('#')
        || id.chars().any(|c| c.is_whitespace())
    {
        bail!("{} must not contain '/', '?', '#' or whitespace", what);
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Channel message reactions (om-je-parity)
// ---------------------------------------------------------------------------
// Graph setReaction/unsetReaction on channel messages — the channel
// counterpart to the chat-service `.../messages/{id}/reactions` calls,
// which return 404 on channel threads (see om-h1-listdetail-PROOF.md).
// Paths follow the v1.0 docs (`/teams/{id}/channels/{id}/messages/{id}/
// (set|unset)Reaction`, `reactionType` sent as unicode); delegated
// `ChannelMessage.Send` rides the first-party client id.
//
// NOT YET VERIFIED LIVE: auth was down when this landed (expired
// refresh token), so the unicode-vs-named `reactionType` form, the
// scope grant, and the Graph-vs-chat-service message id mapping still
// need a signed-in probe before any caller is wired. Receipts stay on
// HOLD: Graph exposes no channel read-receipt API (v1.0 chatMessage
// resource lists no receipt/read-state method), and the chat-service
// consumptionhorizons call 403s on channel threads.

/// `POST .../messages/{id}/setReaction` path. Pure so tests pin it.
pub fn channel_set_reaction_path(team_id: &str, channel_id: &str, message_id: &str) -> String {
    format!(
        "/teams/{}/channels/{}/messages/{}/setReaction",
        team_id.trim(),
        channel_id.trim(),
        message_id.trim()
    )
}

/// `POST .../messages/{id}/unsetReaction` path. Pure so tests pin it.
pub fn channel_unset_reaction_path(
    team_id: &str,
    channel_id: &str,
    message_id: &str,
) -> String {
    format!(
        "/teams/{}/channels/{}/messages/{}/unsetReaction",
        team_id.trim(),
        channel_id.trim(),
        message_id.trim()
    )
}

/// `POST .../messages/{id}/replies/{reply}/setReaction` path.
pub fn channel_reply_set_reaction_path(
    team_id: &str,
    channel_id: &str,
    message_id: &str,
    reply_id: &str,
) -> String {
    format!(
        "/teams/{}/channels/{}/messages/{}/replies/{}/setReaction",
        team_id.trim(),
        channel_id.trim(),
        message_id.trim(),
        reply_id.trim()
    )
}

/// `POST .../messages/{id}/replies/{reply}/unsetReaction` path.
pub fn channel_reply_unset_reaction_path(
    team_id: &str,
    channel_id: &str,
    message_id: &str,
    reply_id: &str,
) -> String {
    format!(
        "/teams/{}/channels/{}/messages/{}/replies/{}/unsetReaction",
        team_id.trim(),
        channel_id.trim(),
        message_id.trim(),
        reply_id.trim()
    )
}

/// `{"reactionType": <emoji>}` body. Graph takes the type as unicode
/// (v1.0 docs example), so the picker emoji itself is sent.
pub fn channel_react_body(emoji: &str) -> serde_json::Value {
    serde_json::json!({ "reactionType": emoji.trim() })
}

/// Set one picker-emoji reaction on a channel message (`reply_id` targets
/// a reply, `None` the root message). Unknown emoji and bad ids are
/// rejected before any network; success is 204 with no body.
pub async fn set_channel_reaction_data(
    client: &TeamsClient,
    team_id: &str,
    channel_id: &str,
    message_id: &str,
    emoji: &str,
    reply_id: Option<&str>,
) -> Result<()> {
    check_id("team_id", team_id)?;
    check_id("channel_id", channel_id)?;
    check_id("message_id", message_id)?;
    if super::chat::reaction_type_for_emoji(emoji.trim()).is_none() {
        bail!("unsupported reaction emoji: {}", emoji);
    }
    let path = match reply_id {
        Some(r) => {
            check_id("reply_id", r)?;
            channel_reply_set_reaction_path(team_id, channel_id, message_id, r)
        }
        None => channel_set_reaction_path(team_id, channel_id, message_id),
    };
    client.graph_post(&path, &channel_react_body(emoji)).await?;
    Ok(())
}

/// Unset one picker-emoji reaction on a channel message. Same arg
/// rules as [`set_channel_reaction_data`].
pub async fn unset_channel_reaction_data(
    client: &TeamsClient,
    team_id: &str,
    channel_id: &str,
    message_id: &str,
    emoji: &str,
    reply_id: Option<&str>,
) -> Result<()> {
    check_id("team_id", team_id)?;
    check_id("channel_id", channel_id)?;
    check_id("message_id", message_id)?;
    if super::chat::reaction_type_for_emoji(emoji.trim()).is_none() {
        bail!("unsupported reaction emoji: {}", emoji);
    }
    let path = match reply_id {
        Some(r) => {
            check_id("reply_id", r)?;
            channel_reply_unset_reaction_path(team_id, channel_id, message_id, r)
        }
        None => channel_unset_reaction_path(team_id, channel_id, message_id),
    };
    client.graph_post(&path, &channel_react_body(emoji)).await?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn react_paths_pin_docs_shapes() {
        assert_eq!(
            channel_set_reaction_path("t1", "19:ch@thread.tacv2", "m1"),
            "/teams/t1/channels/19:ch@thread.tacv2/messages/m1/setReaction"
        );
        assert_eq!(
            channel_unset_reaction_path("t1", "19:ch@thread.tacv2", "m1"),
            "/teams/t1/channels/19:ch@thread.tacv2/messages/m1/unsetReaction"
        );
        assert_eq!(
            channel_reply_set_reaction_path("t1", "19:ch@thread.tacv2", "m1", "r2"),
            "/teams/t1/channels/19:ch@thread.tacv2/messages/m1/replies/r2/setReaction"
        );
        assert_eq!(
            channel_reply_unset_reaction_path("t1", "19:ch@thread.tacv2", "m1", "r2"),
            "/teams/t1/channels/19:ch@thread.tacv2/messages/m1/replies/r2/unsetReaction"
        );
    }

    #[test]
    fn react_body_carries_unicode_reaction_type() {
        assert_eq!(
            channel_react_body("👍"),
            serde_json::json!({ "reactionType": "👍" })
        );
        assert_eq!(
            channel_react_body("  ❤️ "),
            serde_json::json!({ "reactionType": "❤️" })
        );
    }
}
