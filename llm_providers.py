"""
Modular LLM Provider Abstraction Layer

This module provides an abstract interface for LLM providers, allowing
easy integration of multiple providers (Claude, OpenAI, etc.) and models.
"""

from abc import ABC, abstractmethod
from typing import List, Dict, Optional

try:
    from anthropic import Anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False

try:
    import openai
    import httpx
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False


class LLMProvider(ABC):
    """Abstract base class for LLM providers"""
    
    @abstractmethod
    def send_message(self, messages: List[Dict[str, str]], model: str, max_tokens: int = 64000) -> str:
        """
        Send a message to the LLM and return the response.
        
        Args:
            messages: List of message dicts with 'role' and 'content' keys
            model: Model identifier string
            max_tokens: Maximum tokens in response
            
        Returns:
            Response text as string
        """
        pass
    
    @abstractmethod
    def get_available_models(self) -> List[str]:
        """Return list of available model identifiers for this provider"""
        pass
    
    @abstractmethod
    def validate_model(self, model: str) -> bool:
        """Check if a model identifier is valid for this provider"""
        pass


class ClaudeProvider(LLMProvider):
    """Anthropic Claude API provider"""
    
    def __init__(self, api_key: str):
        if not ANTHROPIC_AVAILABLE:
            raise ImportError("anthropic library is not installed. Install it with: pip install anthropic")
        self.client = Anthropic(api_key=api_key)
        # Note: Model identifiers may need to be updated based on actual API availability
        # Check Anthropic API documentation for current model names
        self.available_models = [
            "claude-opus-4-5-20251101",  # Claude Opus 4.5
            "claude-haiku-4-5-20251001",  # Claude Haiku 4.5
            "claude-sonnet-4-5-20250929", # Claude Sonnet 4.5
        ]
    
    def send_message(self, messages: List[Dict[str, str]], model: str, max_tokens: int = 64000) -> str:
        """Send message to Claude API"""
        if not self.validate_model(model):
            raise ValueError(f"Invalid Claude model: {model}")
        
        # Convert messages format for Claude API
        # Claude expects messages in a specific format
        claude_messages = []
        for msg in messages:
            claude_messages.append({
                "role": msg["role"],
                "content": msg["content"]
            })
        
        response = self.client.messages.create(
            model=model,
            max_tokens=max_tokens,
            messages=claude_messages
        )
        
        # Extract text from response
        if response.content and len(response.content) > 0:
            return response.content[0].text
        return ""
    
    def get_available_models(self) -> List[str]:
        """Return list of available Claude models"""
        return self.available_models.copy()
    
    def validate_model(self, model: str) -> bool:
        """Check if model is a valid Claude model"""
        return model in self.available_models


class OpenAIProvider(LLMProvider):
    """OpenAI API provider"""
    
    def __init__(self, api_key: str):
        if not OPENAI_AVAILABLE:
            raise ImportError("openai library is not installed. Install it with: pip install openai")
        self.client = openai.OpenAI(api_key=api_key)
        self.api_key = api_key
        # Create httpx client for custom endpoints like /v1/responses
        self.http_client = httpx.Client(
            base_url="https://api.openai.com",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            },
            timeout=60.0
        )
        # Note: Model identifiers may need to be updated based on actual API availability
        # Check OpenAI API documentation for current model names
        # These are placeholder names - update with actual API model identifiers
        self.available_models = [
            "gpt-5.2-pro",    # GPT-5.2 pro (uses /v1/responses endpoint)
            "gpt-5.2",        # GPT-5.2
            "gpt-5-nano",     # GPT-5 nano
            "gpt-4.1",        # GPT-4.1
        ]
        # Models that use the /v1/responses endpoint instead of /v1/chat/completions
        self.responses_endpoint_models = {
            "gpt-5.2-pro"
        }
    
    def send_message(self, messages: List[Dict[str, str]], model: str, max_tokens: int = 64000) -> str:
        """Send message to OpenAI API"""
        if not self.validate_model(model):
            raise ValueError(f"Invalid OpenAI model: {model}")
        
        # Check if this model uses the /v1/responses endpoint
        if model in self.responses_endpoint_models:
            # Use /v1/responses endpoint for GPT-5.2-Pro
            # Convert messages to a single input string
            input_parts = []
            for msg in messages:
                role = msg.get("role", "user")
                content = msg.get("content", "")
                if role == "user":
                    input_parts.append(content)
                elif role == "assistant":
                    input_parts.append(f"Assistant: {content}")
                elif role == "system":
                    input_parts.append(f"System: {content}")
            
            input_text = "\n\n".join(input_parts)
            
            # Call the /v1/responses endpoint
            response = self.http_client.post(
                "/v1/responses",
                json={
                    "model": model,
                    "input": input_text
                }
            )
            response.raise_for_status()
            
            # Extract text from response - CRITICAL: Never return raw JSON
            try:
                response_data = response.json()
            except Exception as e:
                # If JSON parsing fails, raise error
                raise ValueError(f"Failed to parse response as JSON: {str(e)}")
            
            # Handle the /v1/responses endpoint structure
            # Response can be either a dict or a list
            text_parts = []
            
            if isinstance(response_data, dict):
                # Response is a dict with 'output' or nested structure
                # Check for 'output' field first
                if 'output' in response_data:
                    output = response_data['output']
                    if isinstance(output, str):
                        text_parts.append(output)
                    elif isinstance(output, list):
                        # Output is a list of items
                        for item in output:
                            if isinstance(item, dict):
                                # Look for message type items
                                item_type = item.get('type')
                                if item_type == 'message' and 'content' in item:
                                    content = item.get('content')
                                    if isinstance(content, list):
                                        for content_block in content:
                                            if isinstance(content_block, dict):
                                                block_type = content_block.get('type')
                                                if block_type == 'output_text' and 'text' in content_block:
                                                    text_value = content_block.get('text')
                                                    if text_value is not None:
                                                        text_str = str(text_value).strip()
                                                        if text_str:
                                                            text_parts.append(text_str)
                                                elif 'text' in content_block:
                                                    text_value = content_block.get('text')
                                                    if text_value is not None:
                                                        text_str = str(text_value).strip()
                                                        if text_str:
                                                            text_parts.append(text_str)
                # Check for 'data' field
                elif 'data' in response_data:
                    data = response_data['data']
                    if isinstance(data, list):
                        # Process list of items
                        for item in data:
                            if isinstance(item, dict):
                                item_type = item.get('type')
                                if item_type == 'message' and 'content' in item:
                                    content = item.get('content')
                                    if isinstance(content, list):
                                        for content_block in content:
                                            if isinstance(content_block, dict):
                                                block_type = content_block.get('type')
                                                if block_type == 'output_text' and 'text' in content_block:
                                                    text_value = content_block.get('text')
                                                    if text_value is not None:
                                                        text_str = str(text_value).strip()
                                                        if text_str:
                                                            text_parts.append(text_str)
                                                elif 'text' in content_block:
                                                    text_value = content_block.get('text')
                                                    if text_value is not None:
                                                        text_str = str(text_value).strip()
                                                        if text_str:
                                                            text_parts.append(text_str)
                # Check for direct 'text' or 'response' field
                elif 'text' in response_data:
                    text_value = response_data['text']
                    if text_value is not None:
                        text_str = str(text_value).strip()
                        if text_str:
                            text_parts.append(text_str)
                elif 'response' in response_data:
                    response_value = response_data['response']
                    if isinstance(response_value, str):
                        text_parts.append(response_value)
            
            elif isinstance(response_data, list):
                # Response is a list of objects with 'type' and 'content' fields
                for item in response_data:
                    if not isinstance(item, dict):
                        continue
                    
                    # Look for message type items
                    item_type = item.get('type')
                    if item_type == 'message' and 'content' in item:
                        content = item.get('content')
                        # Content is a list of content blocks
                        if isinstance(content, list):
                            for content_block in content:
                                if isinstance(content_block, dict):
                                    # Look for output_text type
                                    block_type = content_block.get('type')
                                    if block_type == 'output_text' and 'text' in content_block:
                                        text_value = content_block.get('text')
                                        if text_value is not None:
                                            text_str = str(text_value).strip()
                                            if text_str:
                                                text_parts.append(text_str)
                                    # Also check for other text fields (fallback)
                                    elif 'text' in content_block:
                                        text_value = content_block.get('text')
                                        if text_value is not None:
                                            text_str = str(text_value).strip()
                                            if text_str:
                                                text_parts.append(text_str)
            
            # Debug: If no text was extracted, raise error with details
            if not text_parts:
                error_msg = f"No text extracted from response. "
                error_msg += f"Response type: {type(response_data).__name__}. "
                if isinstance(response_data, dict):
                    error_msg += f"Keys: {list(response_data.keys())[:10]}"
                elif isinstance(response_data, list):
                    error_msg += f"Length: {len(response_data)}. "
                    error_msg += f"Item types: {[item.get('type') if isinstance(item, dict) else 'non-dict' for item in response_data[:5]]}"
                raise ValueError(error_msg)
            
            # Return extracted text
            # CRITICAL: Never return str(response_data) or any JSON representation
            result = '\n'.join(text_parts).strip()
            # Final safety check: ensure result is not JSON-like
            if result and len(result) > 0:
                # Only reject if it's clearly a JSON structure (starts with [ or {)
                if result.startswith('[') or result.startswith('{'):
                    # This is definitely JSON, reject it
                    raise ValueError(f"Extracted text looks like JSON (starts with {result[0]}). This should not happen.")
                # Reject only if it contains multiple JSON-like patterns (likely a dict representation)
                # But be very lenient - only reject if we see 4+ patterns together
                json_patterns = ["'id':", '"id":', "'type':", '"type":', "'content':", '"content":', "'status':", '"status":', "'role':", '"role":']
                pattern_count = sum(1 for pattern in json_patterns if pattern in result)
                if pattern_count >= 4:
                    # Too many JSON patterns, likely a dict representation
                    raise ValueError(f"Extracted text contains too many JSON patterns ({pattern_count}). Likely a dict representation.")
                # If we got here, it's valid text - return it
                return result
            
            # If we got here, result is empty
            raise ValueError("Extracted text is empty after processing.")
            
            # Fallback: Handle other possible response structures (dict)
            if isinstance(response_data, dict):
                if 'output' in response_data:
                    return str(response_data['output'])
                elif 'text' in response_data:
                    return str(response_data['text'])
                elif 'response' in response_data:
                    return str(response_data['response'])
                elif 'content' in response_data:
                    content = response_data['content']
                    if isinstance(content, str):
                        return content
                    elif isinstance(content, list):
                        # Try to extract text from content list
                        text_parts = []
                        for item in content:
                            if isinstance(item, dict) and 'text' in item:
                                text_parts.append(item['text'])
                        if text_parts:
                            return '\n'.join(text_parts)
                elif 'data' in response_data:
                    data = response_data['data']
                    if isinstance(data, str):
                        return data
                    elif isinstance(data, dict) and 'output' in data:
                        return str(data['output'])
            
            # If response is a string directly
            if isinstance(response_data, str):
                return response_data
            
            # Last resort: return empty string if structure is unknown (don't return raw JSON)
            return ""
        else:
            # Use chat/completions endpoint for standard chat models
            response = self.client.chat.completions.create(
                model=model,
                messages=messages,
                max_completion_tokens=max_tokens
            )
            
            # Extract text from response
            if response.choices and len(response.choices) > 0:
                return response.choices[0].message.content
            return ""
    
    def get_available_models(self) -> List[str]:
        """Return list of available OpenAI models"""
        return self.available_models.copy()
    
    def validate_model(self, model: str) -> bool:
        """Check if model is a valid OpenAI model"""
        return model in self.available_models


class LLMModelRegistry:
    """Registry for managing LLM models and providers"""
    
    def __init__(self):
        self.providers: Dict[str, LLMProvider] = {}
        self.model_to_provider: Dict[str, str] = {}  # Maps model -> provider name
        self.model_display_names: Dict[str, str] = {}  # Maps model -> display name
    
    def register_provider(self, name: str, provider: LLMProvider):
        """Register a provider and its models"""
        self.providers[name] = provider
        
        # Register all models from this provider
        for model in provider.get_available_models():
            self.model_to_provider[model] = name
            # Generate display name
            display_name = self._generate_display_name(name, model)
            self.model_display_names[model] = display_name
    
    def get_provider_for_model(self, model: str) -> Optional[LLMProvider]:
        """Get the provider instance for a given model"""
        provider_name = self.model_to_provider.get(model)
        if provider_name:
            return self.providers.get(provider_name)
        return None
    
    def get_all_models(self) -> List[str]:
        """Get list of all registered model identifiers"""
        return list(self.model_to_provider.keys())
    
    def get_model_display_name(self, model: str) -> str:
        """Get display name for a model"""
        return self.model_display_names.get(model, model)
    
    def get_models_by_provider(self, provider_name: str) -> List[str]:
        """Get all models for a specific provider"""
        return [model for model, prov in self.model_to_provider.items() if prov == provider_name]
    
    def _generate_display_name(self, provider_name: str, model: str) -> str:
        """Generate a user-friendly display name for a model"""
        # Remove provider prefix and date suffixes
        display = model
        
        # Claude models: claude-opus-4-5-20251101 -> Claude Opus 4.5
        if provider_name.lower() == "claude":
            display = display.replace("claude-", "").replace("-20251101", "")
            parts = display.split("-")
            if len(parts) >= 2:
                model_name = parts[0].capitalize()  # opus -> Opus
                version = " ".join(parts[1:])  # 4-5 -> 4-5
                display = f"Claude {model_name} {version}"
        
        # OpenAI models: gpt-5.2-pro -> GPT-5.2 Pro
        elif provider_name.lower() == "openai":
            display = display.replace("gpt-", "GPT-")
            # Capitalize after hyphens
            parts = display.split("-")
            if len(parts) > 1:
                display = "-".join([parts[0]] + [p.capitalize() for p in parts[1:]])
        
        return display

