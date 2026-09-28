"""Explicit repeat approval used by existing regeneration scenarios."""
import engine

def generate_approved(*args, **kwargs):
    result=engine.generate(*args, **kwargs)
    if result.get('confirmation_required'):
        result=engine.generate(*args, **kwargs, repeat_confirmation=result['repeat_token'])
    return result
