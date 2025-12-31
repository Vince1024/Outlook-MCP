"""List Outlook Rules Tool"""
import json, logging
from utils import get_outlook_application

logger = logging.getLogger(__name__)

def register(mcp):
    @mcp.tool()
    def list_outlook_rules() -> str:
        """List all Outlook rules (mail organization rules)."""
        # Log operation start
        logger.debug("Starting list_outlook_rules operation", extra={
            "operation": "list_outlook_rules"
        })
        
        try:
            outlook = get_outlook_application()
            namespace = outlook.GetNamespace("MAPI")
            rules_collection = namespace.DefaultStore.GetRules()
            rules = []
            for rule in rules_collection:
                try:
                    rule_info = {"name": rule.Name, "enabled": rule.Enabled, "description": "", "conditions": [], "actions": [], "exceptions": []}
                    conditions = rule.Conditions
                    if hasattr(conditions, 'Subject') and conditions.Subject.Enabled:
                        rule_info["conditions"].append(f"Subject contains: {', '.join(conditions.Subject.Text)}")
                    if hasattr(conditions, 'Body') and conditions.Body.Enabled:
                        rule_info["conditions"].append(f"Body contains: {', '.join(conditions.Body.Text)}")
                    if hasattr(conditions, 'From') and conditions.From.Enabled:
                        recipients = [r.Name for r in conditions.From.Recipients]
                        rule_info["conditions"].append(f"From: {', '.join(recipients)}")
                    actions = rule.Actions
                    if hasattr(actions, 'MoveToFolder') and actions.MoveToFolder.Enabled:
                        try:
                            folder_name = actions.MoveToFolder.Folder.Name
                            folder_path = folder_name
                            try:
                                parent = actions.MoveToFolder.Folder.Parent
                                path_parts = [folder_name]
                                while parent and hasattr(parent, 'Name'):
                                    path_parts.insert(0, parent.Name)
                                    parent = parent.Parent if hasattr(parent, 'Parent') else None
                                folder_path = "/".join(path_parts)
                            except Exception:
                                pass
                            rule_info["actions"].append(f"Move to folder: {folder_path}")
                        except Exception:
                            rule_info["actions"].append("Move to folder: (unable to determine)")
                    if hasattr(actions, 'Delete') and actions.Delete.Enabled:
                        rule_info["actions"].append("Delete message")
                    if hasattr(actions, 'MarkAsRead') and actions.MarkAsRead.Enabled:
                        rule_info["actions"].append("Mark as read")
                    desc_parts = []
                    if rule_info["conditions"]:
                        desc_parts.append(f"When: {'; '.join(rule_info['conditions'])}")
                    if rule_info["actions"]:
                        desc_parts.append(f"Then: {'; '.join(rule_info['actions'])}")
                    rule_info["description"] = " | ".join(desc_parts)
                    rules.append(rule_info)
                except Exception:
                    rules.append({"name": rule.Name if hasattr(rule, 'Name') else "Unknown", "enabled": rule.Enabled if hasattr(rule, 'Enabled') else False, "description": "Error parsing rule", "conditions": [], "actions": [], "exceptions": []})
            
            # Log success
            logger.info("Listed Outlook rules successfully", extra={
                "rule_count": len(rules),
                "enabled_rules": sum(1 for r in rules if r.get("enabled", False))
            })
            
            return json.dumps({"success": True, "count": len(rules), "rules": rules}, indent=2)
        except Exception as e:
            logger.error("Failed to list Outlook rules", exc_info=True)
            return json.dumps({"success": False, "error": str(e)})

